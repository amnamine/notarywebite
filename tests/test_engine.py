"""
Unit tests for the core engine:
- split_numbers (LTR runs for numbers like 91/148, 2026/10/05)
- want_dashes and add_dashes (hyphen leader, never doubled)
- empty / dash-only paragraph removal
- detect_box (qism, majmoua, baladiya, wilaya)
- to_grid (rowspan, colspan)
- check_table (subtotals, grand total, uncertain ? cells)
- compute_widths (content based width calculation)
"""
import pytest
from core_engine.xml_helpers import plain_text
from core_engine.ijraa_converter import split_numbers, want_dashes, add_dashes
from core_engine.detect_box import detect_box
from core_engine.tables_builder import to_grid, check_table, compute_widths


def test_split_numbers_in_rtl():
    # Number inside an RTL run should be extracted to its own non-RTL run
    sample = '<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">عقد رقم 91/148 بتاريخ 2026/10/05 مسجل</w:t></w:r>'
    res = split_numbers(sample)
    assert '<w:rtl/>' in res
    assert '<w:r><w:rPr></w:rPr><w:t xml:space="preserve">91/148</w:t></w:r>' in res
    assert '<w:r><w:rPr></w:rPr><w:t xml:space="preserve">2026/10/05</w:t></w:r>' in res


def test_dash_leader_rules():
    # Centered or short paragraphs should not have dashes
    assert not want_dashes("عنوان في الوسط", "center")
    assert not want_dashes("نص قصير.", "both")  # < 60 chars

    # Long text ending with dot or 'أدناه' should get dashes
    long_text = "هذا نص طويل يتجاوز الستين حرفاً لتطبيق قاعدة خطوط النهاية الرسمية في العقد."
    assert want_dashes(long_text, "both")

    # Add dashes: existing trailing dashes stripped, single tab with hyphen leader added
    p_xml = '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>نص تجريبي طويل جداً جداً جداً يتعدى الستين حرفاً للتأكد من الشرط------</w:t></w:r></w:p>'
    added = add_dashes(p_xml, 6690)
    assert 'w:leader="hyphen"' in added
    assert 'w:pos="6690"' in added
    assert added.count('<w:tab/>') == 1


def test_detect_box():
    text = """
    الجمهورية الجزائرية الديمقراطية الشعبية
    مكتب التوثيق
    عقد بيع واقع ببلدية وادي العنب ولاية عنابة،
    ضمن مسح الأراضي، قسم رقم 27 مجموعة ملكية رقم 155
    """
    box = detect_box(text)
    assert '27' in box['qism']
    assert '155' in box['majmoua']
    assert 'وادي العنب' in box['baladiya']
    assert 'عنابة' in box['wilaya']


def test_to_grid_rowspan_colspan():
    # 2 rows table: row 0 has a cell spanning 2 cols, row 1 has 2 single cells
    rows = [
        [{'t': 'Header Spanning', 'cs': 2}],
        [{'t': 'Cell 1'}, {'t': 'Cell 2'}]
    ]
    grid_res = to_grid(rows)
    assert grid_res['ncols'] == 2
    assert len(grid_res['grid']) == 2
    assert grid_res['grid'][0][0]['origin'] is True
    assert grid_res['grid'][0][1]['hcont'] is True
    assert grid_res['grid'][1][0]['origin'] is True
    assert grid_res['grid'][1][1]['origin'] is True


def test_check_table_totals_and_uncertain():
    # Table with uncertain char '?' and a mismatched subtotal
    tb = {
        'headerRows': 1,
        'ncols': 2,
        'grid': [
            [{'ref': {'t': 'البيان', 'cs': 1}, 'origin': True}, {'ref': {'t': 'المساحة', 'cs': 1}, 'origin': True}],
            [{'ref': {'t': 'شقة 1', 'cs': 1}, 'origin': True}, {'ref': {'t': '50,00', 'cs': 1}, 'origin': True}],
            [{'ref': {'t': 'شقة 2?', 'cs': 1}, 'origin': True}, {'ref': {'t': '50,00', 'cs': 1}, 'origin': True}],
            [{'ref': {'t': 'المجموع الجزئي', 'cs': 1}, 'origin': True}, {'ref': {'t': '120,00', 'cs': 1}, 'origin': True}] # 50+50=100 != 120
        ]
    }
    issues = check_table(tb)
    assert any('خانة غير واضحة' in iss for iss in issues)
    assert any('المكتوب 120,00' in iss for iss in issues)


def test_compute_widths():
    tb = {
        'headerRows': 1,
        'ncols': 2,
        'grid': [
            [{'ref': {'t': 'رقم', 'cs': 1}, 'origin': True}, {'ref': {'t': 'تعيين العقار المشاع بالتفصيل', 'cs': 1}, 'origin': True}],
            [{'ref': {'t': '01', 'cs': 1}, 'origin': True}, {'ref': {'t': 'شقة بالطابق الأرضي مع كامل التوابع والمرافق', 'cs': 1}, 'origin': True}]
        ]
    }
    w = compute_widths(tb, 14004)
    assert len(w) == 2
    assert w[1] > w[0]  # The description column should be much wider
    assert sum(w) <= 14004
