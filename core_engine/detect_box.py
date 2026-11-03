"""
Detect cadastral box values from Arabic text.
Direct port of detectBox from core.js.
"""
import re
from typing import Dict, List, Set, Any

STOP_BAL = re.compile(r'^(بتاري|دائرة|ولاي|وولاي|الكائن|رقم|قسم|حي|مسجل|المؤرخ|الموافق)')
STOP_WIL = re.compile(r'^(بتاري|دائرة|بلدي|الكائن|رقم|قسم|حي|المؤرخ|تحت|المسجل)')

def detect_box(text: str) -> Dict[str, str]:
    def uniq(arr: List[str]) -> List[str]:
        seen = set()
        res = []
        for x in arr:
            if x not in seen:
                seen.add(x)
                res.append(x)
        return res

    # qism: /(?:^|[\s(])(?:ال)?قسم\s*(?:رقم)?\s*[:：]?\s*(\d{1,4})(?!\d)/g
    qism_matches = re.finditer(r'(?:^|[\s(])(?:ال)?قسم\s*(?:رقم)?\s*[:：]?\s*(\d{1,4})(?!\d)', text)
    qism_raw = [str(int(m.group(1))).zfill(2) for m in qism_matches]
    qism = uniq(qism_raw)

    # maj: /مجموع[ةه]\s+(?:ال)?ملكي[ةه]\s*(?:رقم)?\s*[:：]?\s*(\d{1,5})/g
    maj_matches = re.finditer(r'مجموع[ةه]\s+(?:ال)?ملكي[ةه]\s*(?:رقم)?\s*[:：]?\s*(\d{1,5})', text)
    maj = uniq([m.group(1) for m in maj_matches])

    def freq(pattern: str, stop: re.Pattern) -> str:
        c = {}
        for m in re.finditer(pattern, text):
            words = []
            extracted = m.group(1)
            for w in re.split(r'\s+', extracted):
                if not w or stop.search(w):
                    break
                words.append(w)
                if len(words) >= 2:
                    break
            k = ' '.join(words)
            k = re.sub(r'^و(?=ال)', '', k)
            
            # window: max(0, m.start() - 250) to m.start() + 250
            start_pos = max(0, m.start() - 250)
            end_pos = min(len(text), m.start() + 250)
            window = text[start_pos:end_pos]
            near = bool(re.search(r'قسم|مجموع[ةه] (?:ال)?ملكي|مسح|المسمى|واقع', window))
            if k:
                c[k] = c.get(k, 0) + (5 if near else 1)
        if not c:
            return ''
        sorted_c = sorted(c.items(), key=lambda item: item[1], reverse=True)
        return sorted_c[0][0]

    bal = freq(r'بلدي[ةه]\s*[:：]?\s*([^،,.:()\-0-9\n]{2,60})', STOP_BAL)
    wil = freq(r'ولاي[ةه]\s*[:：]?\s*([^،,.:()\-0-9\n]{2,40})', STOP_WIL)

    return {
        'qism': ' و '.join(qism[:6]),
        'majmoua': '، '.join(maj[:20]),
        'baladiya': bal,
        'wilaya': wil,
    }
