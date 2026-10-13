# Dựng biên bản họp theo khuôn 3 biên bản mẫu 21/09/2026 (Á Châu, Muối Hà Tĩnh, Tân Hưng).
# Cách dùng:  python build.py <spec.json> <out.docx>
#
# spec.json là 1 list [kind, text], kind thuộc:
#   title  - tiêu đề "BIÊN BẢN HỌP ..."  (căn giữa, đậm)
#   date   - dòng "Ngày dd/mm/yyyy"
#   sec    - tiêu đề mục lớn "I. Thành phần" / "II. Thời gian, địa điểm" / "III. Nội dung chính"
#            (đậm, gạch chân)
#   h      - tiêu đề nhỏ trong mục III, kiểu "Ông X nêu ý kiến:" / "Ý kiến Chủ tịch:" (đậm)
#   b      - gạch đầu dòng cấp 1 (tự thêm "- " ở đầu)
#   b2     - gạch đầu dòng cấp 2, thụt sâu hơn (tự thêm "- " ở đầu)
#   sign   - dòng "Người ghi biên bản" (nghiêng, căn phải)
#   name   - dòng tên người ký, để trống " " nếu chưa có (đậm, căn phải)
#
# Trình bày đã CHỐT theo bản user tự sửa 05/10/2026 (xem memory bien-ban-hop-tu-ghi-am):
# cỡ chữ 12pt toàn văn bản, lề trái/phải 1,19cm, trên 0,76cm, dưới 1,19cm.
# Script này áp luôn các thông số đó khi lưu — KHÔNG cần bước chỉnh tay sau khi build.

import os, sys, json, docx
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.shared import Cm, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, '_template.docx')
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# Trình bày chốt 05/10/2026 — sửa ở đây một lần nếu sau này user chốt thông số khác.
FONT_SIZE_PT = 12
MARGIN_LEFT_CM = 1.19
MARGIN_RIGHT_CM = 1.19
MARGIN_TOP_CM = 0.76
MARGIN_BOTTOM_CM = 1.19

KIND = {
 'title': ('<w:pStyle w:val="Heading1"/><w:spacing w:after="200"/><w:jc w:val="center"/>', '<w:b/><w:bCs/>'),
 'date':  ('<w:spacing w:after="80"/>', ''),
 'sec':   ('<w:spacing w:before="160" w:after="80"/>', '<w:b/><w:bCs/><w:u w:val="single"/>'),
 'b':     ('<w:spacing w:after="60"/><w:ind w:left="360"/>', ''),
 'b2':    ('<w:spacing w:after="60"/><w:ind w:left="720"/>', ''),
 'h':     ('<w:spacing w:after="60"/>', '<w:b/><w:bCs/>'),
 'sign':  ('<w:spacing w:before="400"/><w:jc w:val="right"/>', '<w:i/><w:iCs/>'),
 'name':  ('<w:spacing w:before="400" w:after="0"/><w:jc w:val="right"/>', '<w:b/><w:bCs/>'),
}


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def P(ppr, rpr, text):
    return parse_xml(f'<w:p {W}><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr>'
                      f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def build(spec, out):
    d = docx.Document(TPL)
    body = d.element.body
    for el in list(body):
        if el.tag != qn('w:sectPr'):
            body.remove(el)
    sect = body.find(qn('w:sectPr'))
    for kind, text in spec:
        if kind in ('b', 'b2'):
            text = '- ' + text
        sect.addprevious(P(*KIND[kind], text))

    # Áp trình bày chốt: cỡ chữ + lề.
    for s in d.sections:
        s.left_margin = Cm(MARGIN_LEFT_CM)
        s.right_margin = Cm(MARGIN_RIGHT_CM)
        s.top_margin = Cm(MARGIN_TOP_CM)
        s.bottom_margin = Cm(MARGIN_BOTTOM_CM)
    for p in d.paragraphs:
        for r in p.runs:
            r.font.size = Pt(FONT_SIZE_PT)

    try:
        d.save(out)
    except PermissionError:
        alt = out[:-5] + '_v2.docx' if out.lower().endswith('.docx') else out + '_v2'
        print(f'[canh bao] "{out}" dang bi khoa (co the dang mo trong Word).'
              f' Da luu sang "{alt}" thay the.', file=sys.stderr)
        d.save(alt)
        return alt
    return out


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Dung: python build.py <spec.json> <out.docx>', file=sys.stderr)
        sys.exit(1)
    spec = json.load(open(sys.argv[1], encoding='utf-8'))
    saved_to = build(spec, sys.argv[2])
    print(f'Da luu: {saved_to}')
