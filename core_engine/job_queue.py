"""
Background job worker for Tool B long scans.
Supports:
- Running page by page in a background thread
- Progress tracking (current_page of total_pages)
- Stop flag
- Saving per-page JSON results in data/jobs/<job_id>/
- Resuming or partial download at any time
"""
import json
import os
import threading
from typing import Dict, Any, Optional

from core_engine.models import db, Job, ActivityLog
from core_engine.ai_reader import render_pdf_page, create_image_slices, build_ai_prompt, call_claude_vision
from core_engine.tables_builder import to_grid, build_landscape_doc

_JOB_THREADS: Dict[str, threading.Thread] = {}
_JOB_ABORT_FLAGS: Dict[str, threading.Event] = {}

def get_job_dir(app_data_dir: str, job_id: str) -> str:
    path = os.path.join(app_data_dir, 'jobs', job_id)
    os.makedirs(path, exist_ok=True)
    return path

def stop_job(job_id: str) -> bool:
    if job_id in _JOB_ABORT_FLAGS:
        _JOB_ABORT_FLAGS[job_id].set()
        return True
    return False

def run_tables556_job(app, job_id: str, pdf_path: str, page_start: int, page_end: int, api_key: str, model_name: str):
    with app.app_context():
        job = Job.query.get(job_id)
        if not job:
            return
            
        abort_event = threading.Event()
        _JOB_ABORT_FLAGS[job_id] = abort_event
        job_dir = get_job_dir(app.config['DATA_DIR'], job_id)
        
        job.status = 'processing'
        db.session.commit()

        try:
            with open(pdf_path, 'rb') as f:
                pdf_bytes = f.read()

            import fitz
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            total_pdf_pages = len(doc)
            doc.close()

            actual_start = max(0, page_start)
            actual_end = min(total_pdf_pages, page_end if page_end > 0 else total_pdf_pages)
            
            job.total_pages = actual_end - actual_start
            job.current_page = 0
            db.session.commit()

            last_table = None
            read_ok = 0
            failed = 0

            for idx, p_num in enumerate(range(actual_start, actual_end)):
                if abort_event.is_set():
                    job.status = 'stopped'
                    db.session.commit()
                    return

                page_res_file = os.path.join(job_dir, f'page_{p_num}.json')
                
                # Check if already processed (for resume)
                if os.path.exists(page_res_file):
                    with open(page_res_file, 'r', encoding='utf-8') as f:
                        blocks = json.load(f)
                    read_ok += 1
                else:
                    img_bytes, pdf_text = render_pdf_page(pdf_bytes, p_num)
                    slices = create_image_slices(img_bytes)
                    prompt_text = build_ai_prompt(last_table, len(slices), pdf_text)
                    
                    # Try call Claude with one retry on error
                    blocks = None
                    for attempt in range(2):
                        try:
                            ans = call_claude_vision(api_key, model_name, prompt_text, slices)
                            if ans and isinstance(ans.get('blocks'), list) and not ans.get('unreadable'):
                                blocks = [b for b in ans['blocks'] if b and (b.get('type') == 'text' or (b.get('type') == 'table' and isinstance(b.get('rows'), list)))]
                                break
                        except Exception as e:
                            if attempt == 1:
                                pass

                    if blocks and any(b.get('type') == 'table' for b in blocks):
                        with open(page_res_file, 'w', encoding='utf-8') as f:
                            json.dump(blocks, f, ensure_ascii=False)
                        read_ok += 1
                        for b in blocks:
                            if b.get('type') == 'table':
                                g = to_grid(b['rows'])
                                hr = int(b.get('header_rows') or 0)
                                head = [hr_row for hr_row in b['rows'][:hr]] if hr else (last_table.get('head', []) if last_table else [])
                                last_table = {'cols': g['ncols'], 'head': [[(c.get('t') if isinstance(c, dict) else c) for c in r] for r in head]}
                    else:
                        failed += 1

                job.current_page = idx + 1
                job.read_ok = read_ok
                job.failed_count = failed
                db.session.commit()

            job.status = 'completed'
            db.session.commit()

        except Exception as e:
            job.status = 'failed'
            job.error_message = str(e)
            db.session.commit()
        finally:
            _JOB_ABORT_FLAGS.pop(job_id, None)
            _JOB_THREADS.pop(job_id, None)
