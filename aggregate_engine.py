import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
import re
import copy


def force_font_on_cell(cell, text=None, font_name="가는각진제목체", font_size_pt=10, is_bold=False, alignment=None, space_pt=3):
    """ Clears the cell and forces a clean formatting string. """
    from docx.enum.text import WD_LINE_SPACING
    if text is None: text = cell.text.strip()
        
    p = cell.paragraphs[0]
    p.text = ""
    
    # Completely destroy any hidden list, indent, tab, or spacing formatting in the XML
    pPr = p._element.get_or_add_pPr()
    for tag in ['w:numPr', 'w:ind', 'w:tabs', 'w:spacing', 'w:pStyle']:
        elem = pPr.find(qn(tag))
        if elem is not None:
            pPr.remove(elem)
            
    # Set subtle margins to give breathing room inside cells without causing overflow
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(space_pt)
    p.paragraph_format.space_after = Pt(space_pt)
    
    if alignment is not None:
        p.alignment = alignment

    for extra_p in cell.paragraphs[1:]:
        extra_p._element.getparent().remove(extra_p._element)

    lines = text.split('\n')
    import re
    cleaned_lines = []
    current_indent = ""
    bullet_indent = "  "
    for line in lines:
        line = re.sub(r'\t', '  ', line)
        
        # 괄호 머릿말 검사 (예: [공통]1., [T-타워] 1.)
        prefix_match = re.match(r'^(\s*)([\[\<\(].*?[\]\>\)])(\s*)([0-9가-힣a-zA-Z]\.)', line)
        if prefix_match:
            bracket_part = prefix_match.group(2)
            space_after_bracket = prefix_match.group(3)
            width = sum(2 if ord(c) > 127 else 1 for c in bracket_part) + len(space_after_bracket)
            current_indent = prefix_match.group(1) + (" " * width)
            line = line.rstrip()
            
        elif re.match(r'^\s*[0-9가-힣a-zA-Z]\.', line):
            num_match = re.match(r'^\s*([0-9가-힣a-zA-Z])\.', line)
            if num_match and num_match.group(1) == '1':
                current_indent = ""
                
            line = line.lstrip()
            match = re.match(r'^([0-9가-힣a-zA-Z]\.)\s+(.*)', line)
            if match:
                num = match.group(1)
                content = match.group(2)
                content = re.sub(r' {2,}', ' ', content).strip()
                line = current_indent + num + ' ' + content
                
                # 다음 기호(·, -, 등) 줄이 시작될 위치 계산
                bullet_idx = line.find('·')
                if bullet_idx != -1:
                    prefix_up_to_bullet = line[:bullet_idx]
                    w = sum(2 if ord(c) > 127 else 1 for c in prefix_up_to_bullet)
                    bullet_indent = " " * w
                else:
                    w = sum(2 if ord(c) > 127 else 1 for c in current_indent + num + ' ')
                    bullet_indent = " " * w
            else:
                line = current_indent + line
                w = sum(2 if ord(c) > 127 else 1 for c in current_indent) + 2
                bullet_indent = " " * w
        else:
            line = line.rstrip()
            if re.match(r'^[-※*·]', line.lstrip()):
                if not (line.strip() == '-' or re.match(r'^-\s*\d+$', line)):
                    line = bullet_indent + line.lstrip()
                    
        cleaned_lines.append(line)
        
    for i, line in enumerate(cleaned_lines):
        if i > 0:
            p.add_run('\n')
        run = p.add_run(line)
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.bold = is_bold
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:eastAsia'), font_name)



def flatten_numbering(doc):
    from docx.oxml.ns import qn
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                counter = 1
                for p in cell.paragraphs:
                    pPr = p._element.find(qn('w:pPr'))
                    if pPr is None: continue
                    numPr = pPr.find(qn('w:numPr'))
                    if numPr is None: continue
                    
                    numId_elem = numPr.find(qn('w:numId'))
                    ilvl_elem = numPr.find(qn('w:ilvl'))
                    
                    if numId_elem is None or ilvl_elem is None: continue
                    
                    numId = numId_elem.get(qn('w:val'))
                    ilvl = ilvl_elem.get(qn('w:val'))
                    
                    is_number = False
                    try:
                        numbering_part = doc.part.numbering_part
                        if numbering_part:
                            num = numbering_part.element.find(f'.//w:num[@w:numId="{numId}"]', numbering_part.element.nsmap)
                            if num is not None:
                                abstractNumId = num.find(qn('w:abstractNumId')).get(qn('w:val'))
                                abstractNum = numbering_part.element.find(f'.//w:abstractNum[@w:abstractNumId="{abstractNumId}"]', numbering_part.element.nsmap)
                                if abstractNum is not None:
                                    lvl = abstractNum.find(f'.//w:lvl[@w:ilvl="{ilvl}"]', numbering_part.element.nsmap)
                                    if lvl is not None:
                                        numFmt = lvl.find(qn('w:numFmt')).get(qn('w:val'))
                                        if numFmt in ['decimal', 'lowerLetter', 'upperRoman']:
                                            is_number = True
                    except:
                        pass
                        
                    prefix = f"{counter}. " if is_number else "· "
                    if is_number:
                        counter += 1
                        
                    if len(p.runs) > 0:
                        p.runs[0].text = prefix + p.runs[0].text
                    else:
                        p.add_run(prefix)
                        
                    # Remove numPr so it doesn't get copied as automatic numbering
                    pPr.remove(numPr)


def format_header_cell(cell, override_name=None, force_date=None):
    text = cell.text.strip()
    match = re.search(r'(\d{4}\s*\.\s*\d{1,2}\s*\.\s*\d{1,2})', text)
    if match:
        date_str = match.group(1).replace(" ", "")
        branch_str = text[:match.start()].strip()
    else:
        date_str = ""
        branch_str = text.strip()
        
    if override_name:
        branch_str = override_name
        
    if force_date:
        date_str = force_date
        
    if branch_str.replace(" ", "") == "서린사업부":
        branch_str = "서린 사업부"
        
    # 모든 종류의 앞 공백(특수문자 포함)을 정규식으로 완벽 제거 후 딱 2칸 주입
    branch_str = "  " + re.sub(r'^\s+', '', branch_str)
        
    p = cell.paragraphs[0]
    p.text = ""
    
    # 완전히 초기화: 원본의 들여쓰기 찌꺼기(Indent)가 스페이스바처럼 보이는 현상 방지
    pPr = p._element.get_or_add_pPr()
    for tag in ['w:numPr', 'w:ind', 'w:tabs']:
        elem = pPr.find(qn(tag))
        if elem is not None:
            pPr.remove(elem)
            
    for extra_p in cell.paragraphs[1:]:
        extra_p._element.getparent().remove(extra_p._element)
        
    if branch_str:
        # Check for leading spaces and format without underline
        leading_spaces = len(branch_str) - len(branch_str.lstrip(' '))
        if leading_spaces > 0:
            run_space_prefix = p.add_run(branch_str[:leading_spaces])
            run_space_prefix.font.name = "가는각진제목체"
            run_space_prefix.font.size = Pt(12)
            run_space_prefix.font.underline = False
            run_space_prefix.font.bold = True
            run_space_prefix.font.color.rgb = RGBColor(0, 0, 0)
            
        # Business Unit Name (Underlined and Bold)
        run1 = p.add_run(branch_str.lstrip(' '))
        run1.font.name = "가는각진제목체"
        run1.font.size = Pt(12)
        run1.font.underline = True
        run1.font.bold = True
        run1.font.color.rgb = RGBColor(0, 0, 0)
        rPr1 = run1._element.get_or_add_rPr()
        rPr1.get_or_add_rFonts().set(qn('w:eastAsia'), "가는각진제목체")
        
        # Spaces (Not Underlined)
        run_space = p.add_run("  ")
        run_space.font.name = "가는각진제목체"
        run_space.font.size = Pt(12)
        run_space.font.underline = False
        run_space.font.bold = False
        
    if date_str:
        dm = re.search(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', date_str)
        date_fmt = f"{dm.group(1)}. {int(dm.group(2)):02d}. {int(dm.group(3)):02d}" if dm else date_str
        
        # Date (Bold, Not Underlined)
        run2 = p.add_run(date_fmt)
        run2.font.name = "가는각진제목체"
        run2.font.size = Pt(10)
        run2.font.underline = False
        run2.font.bold = True
        run2.font.color.rgb = RGBColor(0, 0, 0)
        rPr2 = run2._element.get_or_add_rPr()
        rPr2.get_or_add_rFonts().set(qn('w:eastAsia'), "가는각진제목체")

def process_and_merge(template_path, upload_dir, output_path):
    master_doc = Document(template_path)
    
    # Grab Master info
    master_grid = None
    master_sec4_row_xml = None
    if master_doc.tables:
        t0 = master_doc.tables[0]._tbl
        master_grid = t0.find(qn('w:tblGrid'))
        trs = t0.findall(qn('w:tr'))
        if trs:
            master_sec4_row_xml = copy.deepcopy(trs[-1])

    branch_files = {}
    for filename in os.listdir(upload_dir):
        if filename.startswith('~'): continue
        match = re.search(r'^(\d{2})', filename)
        if match:
            idx = int(match.group(1)) - 1
            if 0 <= idx <= 16:
                filepath = os.path.join(upload_dir, filename)
                if idx not in branch_files:
                    branch_files[idx] = filepath
                else:
                    existing = branch_files[idx].lower()
                    current = filepath.lower()
                    if current.endswith('.doc') and not current.endswith('.docx'):
                        branch_files[idx] = filepath
                    elif current.endswith('.docx') and existing.endswith('.docx'):
                        branch_files[idx] = filepath

    import subprocess
    try: subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], capture_output=True)
    except: pass

    global_master_date = None
    if 0 in branch_files:
        try:
            t_doc = Document(branch_files[0])
            flatten_numbering(t_doc)
            if t_doc.tables:
                t_text = t_doc.tables[0].cell(0,0).text
                match = re.search(r'(\d{4}\s*\.\s*\d{1,2}\s*\.\s*\d{1,2})', t_text)
                if match:
                    global_master_date = match.group(1).replace(" ", "")
        except:
            pass

    for idx in range(17):
        if idx not in branch_files:
            continue

        filepath = branch_files[idx]
        try:
            branch_doc = Document(filepath)
            flatten_numbering(branch_doc)
            if not branch_doc.tables:
                continue

            source_tbl = branch_doc.tables[0]._tbl
            target_tbl = master_doc.tables[idx]._tbl
            
            # Get source grid column widths
            source_grid = source_tbl.find(qn('w:tblGrid'))
            scale_factor = 1.0
            source_col_widths = []
            if source_grid is not None:
                for col in source_grid.findall(qn('w:gridCol')):
                    cw = col.get(qn('w:w'))
                    if cw and cw.isdigit():
                        source_col_widths.append(int(cw))
                
                source_total = sum(source_col_widths)
                if source_total > 0:
                    scale_factor = 15168.0 / source_total

            for row in list(target_tbl.findall(qn('w:tr'))):
                target_tbl.remove(row)
                
            if source_grid is not None:
                new_grid = copy.deepcopy(source_grid)
                cols = new_grid.findall(qn('w:gridCol'))
                if scale_factor != 1.0:
                    for col in cols:
                        cw = col.get(qn('w:w'))
                        if cw and cw.isdigit():
                            col.set(qn('w:w'), str(int(int(cw) * scale_factor)))
                
                # FORCE col 0 through 8 width dynamically based on actual header row spans
                header_row = None
                for r in source_tbl.findall(qn('w:tr')):
                    texts = [''.join(tc.itertext()).replace(' ', '').strip() for tc in r.findall(qn('w:tc'))]
                    if any('관리' in t for t in texts) or any('사무' in t for t in texts) or any('사무지원' in t for t in texts):
                        header_row = r
                        break
                        
                if header_row is not None:
                    c_col = 0
                    last_fixed_col_idx = 0
                    header_cells = header_row.findall(qn('w:tc'))
                    num_cells = len(header_cells)
                    for tc_idx, tc in enumerate(header_cells):
                        tcPr = tc.get_or_add_tcPr()
                        gb = tcPr.find(qn('w:gridBefore'))
                        before = int(gb.get(qn('w:val'))) if gb is not None and gb.get(qn('w:val')) else 0
                        c_col += before
                        
                        gs = tcPr.find(qn('w:gridSpan'))
                        span = int(gs.get(qn('w:val'))) if gs is not None and gs.get(qn('w:val')) else 1
                        
                        text = ''.join(tc.itertext()).replace(' ', '').strip()
                        target_w = 0
                        if tc_idx == 0: target_w = 567
                        elif tc_idx == 1: target_w = 992
                        elif 2 <= tc_idx < num_cells - 1:
                            if '차량실' in text:
                                target_w = 1633
                            else:
                                target_w = 862
                        
                        if target_w > 0:
                            part_w = target_w // span
                            remainder = target_w % span
                            for i in range(span):
                                col_idx = c_col + i
                                if col_idx < len(cols):
                                    w_val = part_w + (1 if i < remainder else 0)
                                    cols[col_idx].set(qn('w:w'), str(w_val))
                                    last_fixed_col_idx = max(last_fixed_col_idx, col_idx)
                                    
                        c_col += span
                        
                    remaining_cols = [cols[i] for i in range(last_fixed_col_idx + 1, len(cols))]
                    if remaining_cols:
                        used_w = sum(int(cols[i].get(qn('w:w'))) for i in range(last_fixed_col_idx + 1) if cols[i].get(qn('w:w')))
                        target_rem_w = 15168 - used_w
                        if target_rem_w > 0:
                            current_rem_w = sum(int(c.get(qn('w:w'))) for c in remaining_cols if c.get(qn('w:w')) and c.get(qn('w:w')).isdigit())
                            if current_rem_w > 0:
                                scale_rem = target_rem_w / current_rem_w
                                for c in remaining_cols:
                                    old_w = int(c.get(qn('w:w'))) if c.get(qn('w:w')) and c.get(qn('w:w')).isdigit() else 0
                                    c.set(qn('w:w'), str(int(old_w * scale_rem)))
                            else:
                                remaining_cols[0].set(qn('w:w'), str(target_rem_w))
            else:
                new_grid = None

            for row in source_tbl.findall(qn('w:tr')):
                new_row = copy.deepcopy(row)
                if new_grid is not None:
                    cols = new_grid.findall(qn('w:gridCol'))
                    c_col = 0
                    for tc in new_row.findall(qn('w:tc')):
                        tcPr = tc.get_or_add_tcPr()
                        
                        gb = tcPr.find(qn('w:gridBefore'))
                        before = int(gb.get(qn('w:val'))) if gb is not None and gb.get(qn('w:val')) else 0
                        c_col += before
                        
                        gs = tcPr.find(qn('w:gridSpan'))
                        span = int(gs.get(qn('w:val'))) if gs is not None and gs.get(qn('w:val')) else 1
                        
                        target_w = 0
                        for span_idx in range(span):
                            col_idx = c_col + span_idx
                            if col_idx < len(cols):
                                w_str = cols[col_idx].get(qn('w:w'))
                                if w_str and w_str.isdigit():
                                    target_w += int(w_str)
                                    
                        if target_w > 0:
                            tcW_elements = tcPr.findall(qn('w:tcW'))
                            if not tcW_elements:
                                tcW = parse_xml(f'<w:tcW w:w="{target_w}" w:type="dxa" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                                tcPr.insert(0, tcW)
                            else:
                                for tcW in tcW_elements:
                                    tcW.set(qn('w:w'), str(target_w))
                                    tcW.set(qn('w:type'), 'dxa')
                            
                        c_col += span
                target_tbl.append(new_row)
                
            # Apply correct tblGrid based on branch index
            cg = target_tbl.find(qn('w:tblGrid'))
            if cg is not None: target_tbl.remove(cg)
            
            from docx.oxml import parse_xml
            total_grid_cols = 10
            
            if idx == 12:
                # Gwangmyeong (광명): Use 11-col grid (removed lifeguard)
                total_grid_cols = 11
                gwang_grid_xml = (
                    '<w:tblGrid xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:gridCol w:w="567"/>'
                    '<w:gridCol w:w="993"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="6712"/>'
                    '</w:tblGrid>'
                )
                target_tbl.insert(0, parse_xml(gwang_grid_xml))
            elif idx == 7:
                # Seorin (서린): Use seorin_grid with 1633 dxa (2.88cm) for Col 6
                seorin_grid_xml = (
                    '<w:tblGrid xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:gridCol w:w="567"/>'
                    '<w:gridCol w:w="993"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="1633"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="6803"/>'
                    '</w:tblGrid>'
                )
                target_tbl.insert(0, parse_xml(seorin_grid_xml))
            else:
                # Others: Use clean_grid
                clean_grid_xml = (
                    '<w:tblGrid xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:gridCol w:w="567"/>'
                    '<w:gridCol w:w="993"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="862"/>'
                    '<w:gridCol w:w="7574"/>'
                    '</w:tblGrid>'
                )
                target_tbl.insert(0, parse_xml(clean_grid_xml))
            
            tblPr = target_tbl.find(qn('w:tblPr'))
            if tblPr is not None:
                tblW = tblPr.find(qn('w:tblW'))
                if tblW is not None:
                    tblW.set(qn('w:w'), '15168')
                    tblW.set(qn('w:type'), 'dxa')

            from docx.table import Table
            wrapper_table = Table(target_tbl, master_doc)
            wrapper_table.autofit = False
            
            # --- Pass 1: Ensure Section 4 exists properly ---
            has_sec4 = False
            for row in wrapper_table.rows:
                if not row.cells: continue
                c0_text = row.cells[0].text.replace(" ", "").replace("\n", "").replace("\t", "").replace("\xa0", "")
                if "안전" in c0_text or c0_text == "4" or c0_text == "4.":
                    has_sec4 = True
                    break
            
            if not has_sec4 and master_sec4_row_xml is not None:
                sec4_row = copy.deepcopy(master_sec4_row_xml)
                if 'new_grid' in locals() and new_grid is not None:
                    total_cols = len(new_grid.findall(qn('w:gridCol')))
                    tcs = sec4_row.findall(qn('w:tc'))
                    if tcs and total_cols > 0:
                        used_span = 0
                        used_w = 0
                        for tc in tcs[:-1]:
                            tcPr = tc.find(qn('w:tcPr'))
                            if tcPr is not None:
                                gs = tcPr.find(qn('w:gridSpan'))
                                if gs is not None and gs.get(qn('w:val')) and gs.get(qn('w:val')).isdigit():
                                    used_span += int(gs.get(qn('w:val')))
                                else:
                                    used_span += 1
                                for tcW in tcPr.findall(qn('w:tcW')):
                                    w_val = tcW.get(qn('w:w'))
                                    if w_val and w_val.isdigit():
                                        used_w += int(w_val)
                        
                        last_tc = tcs[-1]
                        last_tcPr = last_tc.get_or_add_tcPr()
                        
                        rem_span = total_cols - used_span
                        if rem_span < 1: rem_span = 1
                        
                        gs = last_tcPr.find(qn('w:gridSpan'))
                        if gs is not None:
                            gs.set(qn('w:val'), str(rem_span))
                        else:
                            gs = parse_xml(f'<w:gridSpan xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{rem_span}"/>')
                            cnfStyle = last_tcPr.find(qn('w:cnfStyle'))
                            if cnfStyle is not None:
                                last_tcPr.insert(last_tcPr.index(cnfStyle) + 1, gs)
                            else:
                                last_tcPr.insert(0, gs)
                            
                        rem_w = 15168 - used_w
                        if rem_w < 100: rem_w = 100
                        for tcW in last_tcPr.findall(qn('w:tcW')):
                            tcW.set(qn('w:w'), str(rem_w))
                            tcW.set(qn('w:type'), 'dxa')
                target_tbl.append(sec4_row)
            
            wrapper_table = Table(target_tbl, master_doc)
                
            # --- Pass 2: Extract text and find Section 4 content cell ---
            sect4_content_cell = None
            extracted_safety_lines = []
            rows_to_delete = []
            seen_cells = set()
            
            current_cat_num = 0

            for row in wrapper_table.rows:
                if not row.cells: continue
                c0_text = row.cells[0].text.replace(" ", "").replace("\n", "")
                c1_text = row.cells[1].text.replace(" ", "").replace("\n", "") if len(row.cells) > 1 else ""

                if "1.인력" in c0_text or c0_text.startswith("1"): current_cat_num = 1
                elif "2.주요" in c0_text or c0_text.startswith("2"): current_cat_num = 2
                elif "3.특이" in c0_text or c0_text.startswith("3") or c0_text == "이상항": current_cat_num = 3
                elif "안전" in c0_text or c0_text == "4" or c0_text == "4.": current_cat_num = 4

                # Identify Section 4 content cell (the one that doesn't hold the title)
                if current_cat_num == 4 and sect4_content_cell is None:
                    for cell in row.cells:
                        c_clean = cell.text.replace(" ", "").replace("\n", "")
                        if c_clean != c0_text and cell not in seen_cells:
                            sect4_content_cell = cell
                            seen_cells.add(cell)
                            break

                if current_cat_num == 3:
                    # Detect rogue Section 4 rows inside Section 3
                    if "안전" in c1_text:
                        for cell in row.cells[2:]:
                            if cell in seen_cells: continue
                            seen_cells.add(cell)
                            txt = cell.text.strip()
                            if txt and txt not in extracted_safety_lines:
                                extracted_safety_lines.append(txt)
                        rows_to_delete.append(row)
                    else:
                        # Scan standard cells for safety lines
                        for cell in row.cells[1:]:
                            if cell in seen_cells: continue
                            seen_cells.add(cell)
                            
                            lines = cell.text.split('\n')
                            kept = []
                            changed = False
                            for ln in lines:
                                if ('안전관리' in ln or '안전문화' in ln) and '안전관리자' not in ln:
                                    if ln.strip() and ln.strip() not in extracted_safety_lines:
                                        extracted_safety_lines.append(ln.strip())
                                    changed = True
                                else:
                                    kept.append(ln)
                            if changed:
                                cell.text = '\n'.join(kept)

            for rogue_row in rows_to_delete:
                try: rogue_row._tr.getparent().remove(rogue_row._tr)
                except: pass

            # --- Pass 3: Append Safety lines to Section 4 ---
            if sect4_content_cell and extracted_safety_lines:
                curr = sect4_content_cell.text.strip()
                to_add = []
                for s in extracted_safety_lines:
                    if s not in curr:
                        to_add.append(s)
                if to_add:
                    if curr: curr += "\n"
                    curr += "\n".join(to_add)
                    sect4_content_cell.text = curr

            # --- Pass 4: Final Format and Overwrite (Headers and Fonts) ---
            seen_cells = set()
            current_cat_num = 0
            
            branch_overrides = {
                8: "SK-P 타워",
                11: "삼화타워",
                12: "광명 U-플래닛 타워 & L7광명 & 아이벡스 컨벤션",
                13: "KT&G 세종공장",
                14: "판교 테크원" 
            }
            
            # Formally fix row 0 height to exactly 1.18 cm
            from docx.shared import Cm
            from docx.enum.table import WD_ROW_HEIGHT_RULE
            
            # Remove direct XML tampering and use robust python-docx built-in method
            wrapper_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            wrapper_table.rows[0].height = Cm(1.18)

            # Structural fix for Gwangmyeong's broken layout
            if len(wrapper_table.rows) > 2:
                tr_1 = wrapper_table.rows[1]._tr
                tr_2 = wrapper_table.rows[2]._tr
                tcs_1 = tr_1.findall(qn('w:tc'))
                tcs_2 = tr_2.findall(qn('w:tc'))
                if len(tcs_1) >= 2 and len(tcs_2) >= 2:
                    tc_1_1 = tcs_1[1]
                    tc_2_1 = tcs_2[1]
                    text_1_1 = "".join(tc_1_1.itertext()).replace(" ", "").strip()
                    text_2_1 = "".join(tc_2_1.itertext()).replace(" ", "").strip()
                    if "유타워" in text_1_1 and "구분" in text_2_1:
                        new_gubun = copy.deepcopy(tc_2_1)
                        tcPr = new_gubun.get_or_add_tcPr()
                        vmerge = tcPr.find(qn('w:vMerge'))
                        if vmerge is not None:
                            vmerge.set(qn('w:val'), 'restart')
                        else:
                            vmerge = parse_xml('<w:vMerge w:val="restart" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                            tcPr.append(vmerge)

                        # Reduce 유타워's span by 1 to make room for 구 분
                        tcPr_utower = tc_1_1.get_or_add_tcPr()
                        gs_utower = tcPr_utower.find(qn('w:gridSpan'))
                        if gs_utower is not None:
                            current_span = int(gs_utower.get(qn('w:val')))
                            if current_span > 1:
                                gs_utower.set(qn('w:val'), str(current_span - 1))

                        new_empty = copy.deepcopy(tc_2_1)
                        for p in new_empty.findall(qn('w:p')):
                            for r in p.findall(qn('w:r')):
                                p.remove(r)
                        tcPr_empty = new_empty.get_or_add_tcPr()
                        vmerge_empty = tcPr_empty.find(qn('w:vMerge'))
                        if vmerge_empty is not None:
                            if qn('w:val') in vmerge_empty.attrib:
                                del vmerge_empty.attrib[qn('w:val')]
                        else:
                            vmerge_empty = parse_xml('<w:vMerge xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                            tcPr_empty.append(vmerge_empty)

                        tr_1.insert(tr_1.index(tc_1_1), new_gubun)
                        idx_2 = tr_2.index(tc_2_1)
                        tr_2.remove(tc_2_1)
                        tr_2.insert(idx_2, new_empty)

            for row_idx, row in enumerate(wrapper_table.rows):
                if not row.cells: continue
                c0_cell = row.cells[0]
                
                # Rule 4: branch renaming and rule 1: underline title
                if row_idx == 0:
                    for cell in row.cells:
                        if cell not in seen_cells:
                            seen_cells.add(cell)
                            format_header_cell(cell, override_name=branch_overrides.get(idx))
                    continue
                
                # Deduplicated Category Hardcoding
                if c0_cell not in seen_cells:
                    c0_text = c0_cell.text.replace(" ", "").replace("\n", "")
                    if "1.인력" in c0_text or (c0_text.startswith("1") and "인력" in c0_text):
                        c0_cell.text = "1. 인력현황"
                        current_cat_num = 1
            # --- Pass 4: Final Format and Overwrite (Headers and Fonts) ---
            seen_cells = set()
            current_cat_num = 0
            
            branch_overrides = {
                8: "SK-P 타워",
                11: "삼화타워",
                12: "광명 U-플래닛 타워 & L7광명 & 아이벡스 컨벤션",
                13: "KT&G 세종공장",
                14: "판교 테크원" 
            }
            
            # Formally fix row 0 height to exactly 1.18 cm
            from docx.shared import Cm
            from docx.enum.table import WD_ROW_HEIGHT_RULE
            
            # Remove direct XML tampering and use robust python-docx built-in method
            wrapper_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            wrapper_table.rows[0].height = Cm(1.18)

            # Structural fix for Gwangmyeong's broken layout
            if len(wrapper_table.rows) > 2:
                tr_1 = wrapper_table.rows[1]._tr
                tr_2 = wrapper_table.rows[2]._tr
                tcs_1 = tr_1.findall(qn('w:tc'))
                tcs_2 = tr_2.findall(qn('w:tc'))
                if len(tcs_1) >= 2 and len(tcs_2) >= 2:
                    tc_1_1 = tcs_1[1]
                    tc_2_1 = tcs_2[1]
                    text_1_1 = "".join(tc_1_1.itertext()).replace(" ", "").strip()
                    text_2_1 = "".join(tc_2_1.itertext()).replace(" ", "").strip()
                    if "유타워" in text_1_1 and "구분" in text_2_1:
                        new_gubun = copy.deepcopy(tc_2_1)
                        tcPr = new_gubun.get_or_add_tcPr()
                        vmerge = tcPr.find(qn('w:vMerge'))
                        if vmerge is not None:
                            vmerge.set(qn('w:val'), 'restart')
                        else:
                            vmerge = parse_xml('<w:vMerge w:val="restart" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                            tcPr.append(vmerge)

                        # Reduce 유타워's span by 1 to make room for 구 분
                        tcPr_utower = tc_1_1.get_or_add_tcPr()
                        gs_utower = tcPr_utower.find(qn('w:gridSpan'))
                        if gs_utower is not None:
                            current_span = int(gs_utower.get(qn('w:val')))
                            if current_span > 1:
                                gs_utower.set(qn('w:val'), str(current_span - 1))

                        new_empty = copy.deepcopy(tc_2_1)
                        for p in new_empty.findall(qn('w:p')):
                            for r in p.findall(qn('w:r')):
                                p.remove(r)
                        tcPr_empty = new_empty.get_or_add_tcPr()
                        vmerge_empty = tcPr_empty.find(qn('w:vMerge'))
                        if vmerge_empty is not None:
                            if qn('w:val') in vmerge_empty.attrib:
                                del vmerge_empty.attrib[qn('w:val')]
                        else:
                            vmerge_empty = parse_xml('<w:vMerge xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                            tcPr_empty.append(vmerge_empty)

                        tr_1.insert(tr_1.index(tc_1_1), new_gubun)
                        idx_2 = tr_2.index(tc_2_1)
                        tr_2.remove(tc_2_1)
                        tr_2.insert(idx_2, new_empty)

            for row_idx, row in enumerate(wrapper_table.rows):
                if not row.cells: continue
                
                # 타이틀 행(row_idx == 0)은 원본의 예쁜 여백과 고정 높이를 그대로 보존하고,
                # 그 외의 일반 데이터 행들만 강제 높이 고정(trHeight)을 파괴하여 AutoFit 유도
                if row_idx > 0:
                    trPr = row._tr.get_or_add_trPr()
                    for h in trPr.findall(qn('w:trHeight')):
                        trPr.remove(h)
                    
                c0_cell = row.cells[0]
                
                # Rule 4: branch renaming and rule 1: underline title
                if row_idx == 0:
                    for cell in row.cells:
                        if cell not in seen_cells:
                            seen_cells.add(cell)
                            format_header_cell(cell, override_name=branch_overrides.get(idx), force_date=global_master_date)
                    continue
                
                # Deduplicated Category Hardcoding
                if c0_cell not in seen_cells:
                    c0_text = c0_cell.text.replace(" ", "").replace("\n", "")
                    if "1.인력" in c0_text or (c0_text.startswith("1") and "인력" in c0_text):
                        c0_cell.text = "1. 인력현황"
                        current_cat_num = 1
                    elif "2.주요" in c0_text or (c0_text.startswith("2") and "주요" in c0_text):
                        c0_cell.text = "2. 주요추진업무"
                        current_cat_num = 2
                    elif "3.특이" in c0_text or (c0_text.startswith("3") and "특이" in c0_text) or c0_text in ["이상항", "3", "3."]:
                        c0_cell.text = "3. 특이사항"
                        current_cat_num = 3
                    elif "안전" in c0_text or c0_text == "4" or c0_text == "4.":
                        if len(row.cells) > 1 and "안전문화" in row.cells[1].text.replace(" ", ""):
                            c0_cell.merge(row.cells[1])
                        c0_cell.text = "4. 안전문화활동"
                        current_cat_num = 4

                if current_cat_num == 1 and len(row.cells) > 1:
                    c1_text = row.cells[1].text.replace(' ', '').replace('\n', '')
                    if '구분' in c1_text:
                        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                        row.height = Cm(0.56)
                    elif '정원' in c1_text:
                        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                        row.height = Cm(0.56)
                    elif '현원' in c1_text:
                        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                        row.height = Cm(0.56)
                    elif '증감' in c1_text:
                        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                        row.height = Cm(0.56)

                seen_cells = set()
                row_text_raw = ''.join(row._tr.itertext()).replace(' ', '')
                is_personnel_header = any(x in row_text_raw for x in ['관리', '시설', '보안']) and not any(x in row_text_raw for x in ['정원', '현원', '증감'])
                topic_count = 0
                replace_next_with_content = False
                for cell_idx, cell in enumerate(row.cells):
                    # Remove background shading unconditionally for the safety row
                    if current_cat_num == 4:
                        tcPr = cell._tc.get_or_add_tcPr()
                        for shd in tcPr.findall(qn('w:shd')):
                            tcPr.remove(shd)

                    if cell in seen_cells: continue
                    seen_cells.add(cell)

                    text = cell.text.strip()
                    if current_cat_num == 1 and not is_personnel_header:
                        if idx != 13:
                            if text == '' or text == '-':
                                text = '0'
                        else:
                            if cell_idx in [5, 6]:
                                text = ''

                    if '차량실' in text and '주차' in text and '관제' in text:
                        text = text.replace(' ', '')

                    text_no_space = text.replace(' ', '')
                    text_no_newline = text_no_space.replace('\n', '')
                    align = None
                    
                    if current_cat_num == 1 and cell_idx < len(row.cells) - 1:
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        align = WD_ALIGN_PARAGRAPH.CENTER
                    
                    if replace_next_with_content:
                        text = '내       용'
                        text_no_space = '내용'
                        text_no_newline = '내용'
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        align = WD_ALIGN_PARAGRAPH.CENTER
                        replace_next_with_content = False
                    
                    if text_no_newline in ['주제', '주체']:
                        topic_count += 1
                        if topic_count == 2:
                            text = '구분'
                            text_no_space = '구분'
                            text_no_newline = '구분'
                            from docx.enum.text import WD_ALIGN_PARAGRAPH
                            align = WD_ALIGN_PARAGRAPH.CENTER
                            replace_next_with_content = True
                            
                            # Separate from upper column grid
                            if cell_idx + 1 < len(row.cells):
                                next_c = row.cells[cell_idx + 1]
                                c_tcPr = cell._tc.get_or_add_tcPr()
                                c_gs = c_tcPr.find(qn('w:gridSpan'))
                                if c_gs is None:
                                    from docx.oxml import OxmlElement
                                    c_gs = OxmlElement('w:gridSpan')
                                    c_gs.set(qn('w:val'), '2')
                                    c_tcPr.append(c_gs)
                                n_tcPr = next_c._tc.get_or_add_tcPr()
                                n_gs = n_tcPr.find(qn('w:gridSpan'))
                                if n_gs is not None and int(n_gs.get(qn('w:val'), 1)) > 1:
                                    n_gs.set(qn('w:val'), str(int(n_gs.get(qn('w:val'))) - 1))
                            
                    if text_no_newline in ['행복지원팀', '역량육성팀']:
                        from docx.shared import Cm
                        cell.width = Cm(2.01)
                        
                        # Separate from upper column grid
                        if cell_idx + 1 < len(row.cells):
                            next_c = row.cells[cell_idx + 1]
                            c_tcPr = cell._tc.get_or_add_tcPr()
                            c_gs = c_tcPr.find(qn('w:gridSpan'))
                            if c_gs is None:
                                from docx.oxml import OxmlElement
                                c_gs = OxmlElement('w:gridSpan')
                                c_gs.set(qn('w:val'), '2')
                                c_tcPr.append(c_gs)
                            n_tcPr = next_c._tc.get_or_add_tcPr()
                            n_gs = n_tcPr.find(qn('w:gridSpan'))
                            if n_gs is not None and int(n_gs.get(qn('w:val'), 1)) > 1:
                                n_gs.set(qn('w:val'), str(int(n_gs.get(qn('w:val'))) - 1))
                        
                    if '이슈(기타)' in text_no_space:
                        text = '이슈\n(기타)'
                    elif '직원동향' in text_no_space:
                        text = '직원\n동향'
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        align = WD_ALIGN_PARAGRAPH.CENTER
                    elif text_no_newline == '고객사':
                        text = '고객사'
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        align = WD_ALIGN_PARAGRAPH.CENTER
                    elif text_no_newline == '우리회사':
                        text = '우리회사'
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        align = WD_ALIGN_PARAGRAPH.CENTER
                    elif text_no_newline == '내용':
                        # T타워 양식에 맞춰 '내'와 '용' 사이에 스페이스바 7칸 강제 주입 및 가운데 정렬
                        text = '내       용'
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        align = WD_ALIGN_PARAGRAPH.CENTER
                    elif text_no_space in ['유타워', 'L7광명', '아이벡스']:
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        align = WD_ALIGN_PARAGRAPH.CENTER
                    elif current_cat_num == 1 and is_personnel_header and ('행정' in text or '사무' in text or '서무' in text):
                        text = '사무지원'
                    elif current_cat_num == 1 and is_personnel_header and '안내' in text:
                        text = 'CS'
                        
                    # 1. 인력현황의 마지막 열(결원/대책 칸)이 빈칸인 경우 기본 텍스트 주입
                    if current_cat_num == 1 and is_personnel_header and cell_idx == len(row.cells) - 1:
                        if not text_no_newline:
                            text = '1. 결원 : 0명\n2. 대책 : '

                    # 4. 안전문화활동 내용 중 불필요한 '안전관리' 헤더 텍스트 제거
                    if current_cat_num == 4:
                        text = re.sub(r'(?m)^\s*안전관리\s*\n?', '', text)

                    # Calculate dynamic margin for "4. 안전문화활동"
                    custom_space_pt = 8 if current_cat_num == 4 else 3

                    # Apply text cleanly without any bold (Rule 2, 3)
                    force_font_on_cell(cell, text, font_name="가는각진제목체", font_size_pt=10, is_bold=False, alignment=align, space_pt=custom_space_pt)

            # --- Pass 3.4: KT&G Sejong (idx == 13) Border Fix ---
            if idx == 13:
                for tr in wrapper_table._tbl.findall(qn('w:tr')):
                    for tcPr in tr.findall(f".//{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}tcPr"):
                        tcBorders = tcPr.find(qn('w:tcBorders'))
                        if tcBorders is not None:
                            for border in tcBorders:
                                val = border.get(qn('w:val'))
                                if val in ['double', 'dotted', 'dashed', 'dotDash', 'dotDotDash', 'dashSmallGap', 'dashDotStroked']:
                                    border.set(qn('w:val'), 'single')

            # --- Pass 3.5: Gwangmyeong Custom Merge ---
            if idx == 12:
                from docx.table import _Cell
                for tr in wrapper_table._tbl.findall(qn('w:tr')):
                    # 점선/파선을 모두 실선으로 변환
                    for tcPr in tr.findall(f".//{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}tcPr"):
                        tcBorders = tcPr.find(qn('w:tcBorders'))
                        if tcBorders is not None:
                            for border in tcBorders:
                                val = border.get(qn('w:val'))
                                if val in ['dotted', 'dashed', 'dotDash', 'dotDotDash', 'dashSmallGap', 'dashDotStroked']:
                                    border.set(qn('w:val'), 'single')
                                    
                    tcs = tr.findall(qn('w:tc'))
                    if len(tcs) == 12:
                        # Extract text safely without invoking row.cells layout engine
                        txt8 = _Cell(tcs[8], wrapper_table).text.strip().replace(chr(10), '').replace(' ', '')
                        txt9 = _Cell(tcs[9], wrapper_table).text.strip().replace(chr(10), '').replace(' ', '')
                        
                        if txt8.isdigit() and txt9.isdigit():
                            new_val = str(int(txt8) + int(txt9))
                        else:
                            new_val = "보안" # Header
                        
                        # Set text safely with font and centering XML to preserve styling
                        tcs[8].clear_content()
                        p = etree.SubElement(tcs[8], qn('w:p'))
                        pPr = etree.SubElement(p, qn('w:pPr'))
                        jc = etree.SubElement(pPr, qn('w:jc'))
                        jc.set(qn('w:val'), 'center')
                        
                        r = etree.SubElement(p, qn('w:r'))
                        rPr = etree.SubElement(r, qn('w:rPr'))
                        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
                        rFonts.set(qn('w:ascii'), '가는각진제목체')
                        rFonts.set(qn('w:eastAsia'), '가는각진제목체')
                        rFonts.set(qn('w:hAnsi'), '가는각진제목체')
                        
                        sz = etree.SubElement(rPr, qn('w:sz'))
                        sz.set(qn('w:val'), '20')  # 10pt
                        
                        color = etree.SubElement(rPr, qn('w:color'))
                        color.set(qn('w:val'), '000000')
                        
                        t_el = etree.SubElement(r, qn('w:t'))
                        t_el.text = new_val
                        
                        # Remove the col9 w:tc
                        tr.remove(tcs[9])
            # --- Pass 3.6: Dynamic Row Splitting ---
            from docx.table import _Cell
            from docx.oxml import OxmlElement

            def set_vmerge(tc, val):
                tcPr = tc.get_or_add_tcPr()
                vMerge = tcPr.find(qn('w:vMerge'))
                if vMerge is None:
                    vMerge = OxmlElement('w:vMerge')
                    tcPr.append(vMerge)
                if val:
                    vMerge.set(qn('w:val'), val)
                else:
                    if qn('w:val') in vMerge.attrib:
                        del vMerge.attrib[qn('w:val')]

            # --- Jungbu (idx == 2) logic removed to respect original document rows ---

            # --- Busan (idx == 3) and Daegu (idx == 4) ---
            if idx in [3, 4, 5, 6]:
                for row in wrapper_table.rows:
                    cells = row.cells
                    if len(cells) > 1 and '이슈' in cells[1].text and '기타' in cells[1].text:
                        tr_orig = row._tr
                        tcs_orig = tr_orig.findall(qn('w:tc'))
                        
                        if len(tcs_orig) == 3:
                            trPr = tr_orig.find(qn('w:trPr'))
                            if trPr is not None:
                                for th in trPr.findall(qn('w:trHeight')): trPr.remove(th)
                                
                            content_text = _Cell(tcs_orig[2], wrapper_table).text.strip()
                            parsed_lines = []
                            
                            if idx in [3, 4]:
                                lines = content_text.split('\n')
                                current_prefix = None
                                current_suffix = []

                                for line in lines:
                                    line = line.strip()
                                    if not line: continue
                                    if ':' in line and len(line.split(':', 1)[0]) <= 10:
                                        if current_prefix:
                                            parsed_lines.append((current_prefix, '\n'.join(current_suffix)))
                                        parts = line.split(':', 1)
                                        prefix = parts[0].split('.')[-1].strip()
                                        current_prefix = prefix
                                        current_suffix = [parts[1].strip()]
                                    else:
                                        if current_prefix:
                                            current_suffix.append(line)

                                if current_prefix:
                                    parsed_lines.append((current_prefix, '\n'.join(current_suffix)))
                                    
                            elif idx in [5, 6]:
                                if idx == 5:
                                    targets = ["공통", "우산", "송정", "전주", "현대사업장 및 SK 넥실리스 사업장"]
                                else:
                                    targets = ["공통", "제주 사옥", "오리온 제주 용암수", "제주 쉬멍"]

                                lines = content_text.split('\n')
                                extracted = {}
                                current_key = None
                                
                                for line in lines:
                                    sline = line.strip()
                                    if not sline: continue
                                    
                                    m = re.match(r'^[0-9]+\.\s*(.*)', sline)
                                    is_header = False
                                    if m:
                                        header_text = m.group(1).strip()
                                        header_text_clean = header_text.replace(" ", "")
                                        for t in targets:
                                            t_clean = t.replace(" ", "")
                                            if t_clean in header_text_clean or header_text_clean in t_clean:
                                                current_key = t
                                                if current_key not in extracted: extracted[current_key] = []
                                                is_header = True
                                                break
                                            elif "현대" in t and "넥실리스" in t:
                                                if "현대" in header_text_clean and "넥실리스" in header_text_clean:
                                                    current_key = t
                                                    if current_key not in extracted: extracted[current_key] = []
                                                    is_header = True
                                                    break
                                            elif t == "제주 쉬멍" and "쉬멍" in header_text:
                                                current_key = t
                                                if current_key not in extracted: extracted[current_key] = []
                                                is_header = True
                                                break
                                    
                                    if not is_header:
                                        if current_key:
                                            extracted[current_key].append(sline)
                                
                                for t in targets:
                                    if t in extracted and extracted[t]:
                                        parsed_lines.append((t, '\n'.join(extracted[t])))
                                    else:
                                        parsed_lines.append((t, '특이사항 없음'))
                            
                            if not parsed_lines:
                                break
                                
                            def create_cell_with_span(span, text, font_name="가는각진제목체"):
                                tc = OxmlElement('w:tc')
                                tcPr = OxmlElement('w:tcPr')
                                tc.append(tcPr)
                                gs = OxmlElement('w:gridSpan')
                                gs.set(qn('w:val'), str(span))
                                tcPr.append(gs)
                                tcBorders = OxmlElement('w:tcBorders')
                                for b_name in ['top', 'left', 'bottom', 'right']:
                                    b = OxmlElement(f'w:{b_name}')
                                    b.set(qn('w:val'), 'single')
                                    b.set(qn('w:sz'), '4')
                                    b.set(qn('w:space'), '0')
                                    b.set(qn('w:color'), 'auto')
                                    tcBorders.append(b)
                                tcPr.append(tcBorders)
                                vAlign = OxmlElement('w:vAlign')
                                vAlign.set(qn('w:val'), 'center')
                                tcPr.append(vAlign)
                                p = OxmlElement('w:p')
                                pPr = OxmlElement('w:pPr')
                                if span == 2:
                                    jc = OxmlElement('w:jc')
                                    jc.set(qn('w:val'), 'center')
                                    pPr.append(jc)
                                p.append(pPr)
                                tc.append(p)
                                force_font_on_cell(_Cell(tc, wrapper_table), text, font_name=font_name, font_size_pt=10, is_bold=False)
                                return tc
                            
                            set_vmerge(tcs_orig[1], 'restart')
                            tr_orig.remove(tcs_orig[2])
                            
                            branch_tc = create_cell_with_span(2, parsed_lines[0][0], font_name="가는각진제목체")
                            tr_orig.append(branch_tc)
                            content_tc = create_cell_with_span(6, parsed_lines[0][1])
                            tr_orig.append(content_tc)
                            
                            for i in range(1, len(parsed_lines)):
                                new_tr = copy.deepcopy(tr_orig)
                                new_tcs = new_tr.findall(qn('w:tc'))
                                
                                set_vmerge(new_tcs[0], 'continue')
                                set_vmerge(new_tcs[1], 'continue')
                                
                                _Cell(new_tcs[0], wrapper_table).text = ""
                                _Cell(new_tcs[1], wrapper_table).text = ""
                                
                                force_font_on_cell(_Cell(new_tcs[2], wrapper_table), parsed_lines[i][0], font_name="가는각진제목체", font_size_pt=10, is_bold=False)
                                force_font_on_cell(_Cell(new_tcs[3], wrapper_table), parsed_lines[i][1], font_name="가는각진제목체", font_size_pt=10, is_bold=False)
                                
                                tr_orig.getparent().insert(tr_orig.getparent().index(tr_orig) + i, new_tr)
                        break

            # --- Gongsa Section Split for Specific Branches (idx in [7, 8, 10, 11, 13]) ---
            if idx in [7, 8, 10, 11, 13]:
                site_names = {
                    7: "서린",
                    8: "판교",
                    10: "종로타워",
                    11: "삼화타워",
                    13: "KT&G"
                }
                site_name = site_names[idx]
                
                for row in wrapper_table.rows:
                    # Resolve unique cells to handle different grids (10-col vs 12-col)
                    unique_cells = []
                    for cell in row.cells:
                        if not unique_cells or cell._tc != unique_cells[-1]._tc:
                            unique_cells.append(cell)
                            
                    if len(unique_cells) == 4:
                        c1_clean = re.sub(r'\s+', '', unique_cells[1].text)
                        c2_clean = re.sub(r'\s+', '', unique_cells[2].text)
                        
                        if c1_clean == '공사' and c2_clean in ['주체', '고객사', '우리회사']:
                            tr = row._tr
                            tc_to_copy = unique_cells[2]._tc
                            new_tc = copy.deepcopy(tc_to_copy)
                            val_to_set = "구분" if c2_clean == '주체' else site_name
                            
                            tr.insert(tr.index(tc_to_copy) + 1, new_tc)
                            cell_obj = _Cell(new_tc, wrapper_table)
                            force_font_on_cell(cell_obj, val_to_set)

            # 4. Enforce Unified Widths and Exact 10-Column GridSpan (Cell-Count Based)
            from lxml import etree
            
            for row in wrapper_table.rows:
                row_xml = row._tr
                tcs = row_xml.findall(qn('w:tc'))
                if not tcs:
                    continue
                
                num_cells = len(tcs)
                used_span = 0
                
                for i, tc in enumerate(tcs):
                    tcPr = tc.get_or_add_tcPr()
                    
                    # 1. Clear old gridSpan to prevent overflow
                    for old_gs in tcPr.findall(qn('w:gridSpan')):
                        tcPr.remove(old_gs)
                        
                    # 2. Determine target width and span based on exact grid columns
                    span = 1
                    
                    if num_cells == 4:
                        if i == 2 and not (idx == 12 and "인력현황" in ''.join(row_xml.itertext())):
                            span = 2
                        elif idx == 12 and i == 2 and "인력현황" in ''.join(row_xml.itertext()):
                            span = 8
                    elif num_cells == 2:
                        if i == 0: span = 2
                    
                    # Ensure sum of spans is exactly total_grid_cols
                    if i == num_cells - 1:
                        span = total_grid_cols - used_span
                        if span < 1: span = 1
                        
                    # Calculate EXACT width from grid column definitions to prevent any proportional resizing
                    grid_col_widths = [567, 993, 862, 862, 862, 862, 862, 862, 862, 7574]
                    if idx == 7: # Seorin
                        grid_col_widths = [567, 993, 862, 862, 862, 862, 1633, 862, 862, 6803]
                    elif idx == 12: # Gwangmyeong
                        grid_col_widths = [567, 993, 862, 862, 862, 862, 862, 862, 862, 862, 6712]
                        
                    # Ensure we don't index out of bounds
                    end_col = min(used_span + span, total_grid_cols)
                    target_w = sum(grid_col_widths[used_span : end_col])
                    
                    used_span += span
                    
                    if span > 1:
                        gs = etree.SubElement(tcPr, qn('w:gridSpan'))
                        gs.set(qn('w:val'), str(span))
                    
                    if target_w:
                        for old in tcPr.findall(qn('w:tcW')): tcPr.remove(old)
                        tcW = etree.SubElement(tcPr, qn('w:tcW'))
                        tcW.set(qn('w:w'), str(target_w))
                        tcW.set(qn('w:type'), 'dxa')
                        
                    # Fix XML element sequence to prevent Word corruption
                    order = [
                        'w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge',
                        'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMargin', 'w:textDirection',
                        'w:tcFitText', 'w:vAlign', 'w:hideMark', 'w:headers', 'w:cellIns',
                        'w:cellDel', 'w:cellMerge'
                    ]
                    def get_order(el):
                        for i, name in enumerate(order):
                            if el.tag == qn(name): return i
                        return 999
                    elements = list(tcPr)
                    elements.sort(key=get_order)
                    for el in elements:
                        tcPr.append(el)

        except Exception as e:
            import traceback
            traceback.print_exc()

    # --- Pass 5: Clean empty paras and ensure page breaks (No blank page at end) ---
    body = master_doc.element.body
    for p in list(body.findall(qn('w:p'))):
        if not p.xpath('.//w:t'):
            p.getparent().remove(p)

    tbl_elements = body.findall(qn('w:tbl'))
    for i, tbl in enumerate(tbl_elements[:-1]):
        new_p_xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:br w:type="page"/></w:r></w:p>'
        new_p_elem = etree.fromstring(new_p_xml)
        tbl.addnext(new_p_elem)

    master_doc.save(output_path)
    return master_doc
