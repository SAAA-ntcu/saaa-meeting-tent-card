#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Meeting Tent Card Engine (會議桌牌生成與維護核心引擎)
======================================================
Automates the creation, inspection, modification, and export of standard
Taiwan A4 4-fold triangular meeting tent cards (會議桌牌 / 立體桌牌) in Word (.docx).

Features:
- Dual-Rendering synchronization: DrawingML (0° & 180° rot) + VML Fallback (0° & flip:x y).
- Intelligent auto-fit typography: Adaptive font size and boundary-aware unit line wrapping.
- Exact print-ready A4 4-fold layout: 7.43cm + 8.45cm + 8.45cm + 1.97cm.
- Zero-spill isolation: Guaranteed 1 card per page via section break management.
- Multi-source input: Excel (.xlsx), CSV, TSV, JSON, raw text.
"""

import os
import sys
import copy
import json
import zipfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any

# Ensure stdout/stderr utf-8 on Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# OpenXML Namespaces
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
V_NS = 'urn:schemas-microsoft-com:vml'
W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
WPG_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
W10_NS = 'urn:schemas-microsoft-com:office:word'

def register_namespaces():
    ET.register_namespace('w', W_NS)
    ET.register_namespace('wps', WPS_NS)
    ET.register_namespace('v', V_NS)
    ET.register_namespace('w14', W14_NS)
    ET.register_namespace('mc', MC_NS)
    ET.register_namespace('wp', WP_NS)
    ET.register_namespace('a', A_NS)
    ET.register_namespace('pic', PIC_NS)
    ET.register_namespace('wpg', WPG_NS)
    ET.register_namespace('w10', W10_NS)

register_namespaces()

DEFAULT_TEMPLATE_PATH = Path(__file__).resolve().parent.parent / 'templates' / 'template_a4_tent.docx'


class TentCardEngine:
    def __init__(self, template_path: Optional[str] = None):
        tpl = Path(template_path) if template_path else DEFAULT_TEMPLATE_PATH
        if not tpl.exists():
            raise FileNotFoundError(f"Template docx not found at: {tpl}")
        self.template_path = tpl
        self._load_template()

    def _load_template(self):
        """Loads and parses the master template docx into XML components."""
        with zipfile.ZipFile(self.template_path) as z:
            self.file_map = {name: z.read(name) for name in z.namelist()}
        
        self.root = ET.fromstring(self.file_map['word/document.xml'])
        body = self.root.find(f'{{{W_NS}}}body')
        children = list(body)
        
        if len(children) < 4:
            raise ValueError("Template structure invalid: expected at least table, section break, final p, and sectPr.")
        
        # Prototype elements
        self.proto_tbl = copy.deepcopy(children[0])
        self.proto_p = copy.deepcopy(children[1])
        self.final_p = copy.deepcopy(children[-2])
        self.final_sectPr = copy.deepcopy(children[-1])

    @staticmethod
    def split_unit_text(unit_str: str) -> List[str]:
        """Intelligently splits affiliation/unit text into 1 or 2 lines."""
        if not unit_str:
            return []
        
        # Respect user explicit line breaks
        if '\n' in unit_str:
            return [l.strip() for l in unit_str.split('\n') if l.strip()]
        if ' / ' in unit_str or ' | ' in unit_str:
            delim = ' / ' if ' / ' in unit_str else ' | '
            return [l.strip() for l in unit_str.split(delim) if l.strip()]
        if '  ' in unit_str:
            return [l.strip() for l in unit_str.split('  ') if l.strip()]
        
        # If short enough, keep 1 line
        if len(unit_str) <= 12:
            return [unit_str]
        
        # If contains single space, check if splitting makes sense
        if ' ' in unit_str:
            parts = [p.strip() for p in unit_str.split(' ') if p.strip()]
            if len(parts) == 2:
                return parts
        
        # Common administrative & academic boundary splitting
        split_keywords = [
            '科技大學', '教育大學', '師範大學', '醫學大學', '大學',
            '學院', '高級中學', '專科學校', '國民中學', '國民小學',
            '政府教育局', '政府教育處', '國小', '國中', '高中', '學校'
        ]
        for kw in split_keywords:
            if kw in unit_str:
                idx = unit_str.index(kw) + len(kw)
                if 3 <= idx < len(unit_str):
                    return [unit_str[:idx], unit_str[idx:]]
        
        # Fallback: split at center
        mid = len(unit_str) // 2
        return [unit_str[:mid], unit_str[mid:]]

    @staticmethod
    def calculate_font_sizes(u_lines: List[str], full_name_title: str, 
                             force_u_sz: Optional[int] = None, 
                             force_nt_sz: Optional[int] = None) -> Tuple[str, str]:
        """Calculates half-pt font sizes (e.g. 76 = 38pt) for Unit and Name/Title."""
        if force_u_sz:
            u_sz = str(force_u_sz * 2)
        else:
            max_u = max(len(l) for l in u_lines) if u_lines else 0
            if max_u > 16:
                u_sz = '52' # 26pt
            elif max_u > 14:
                u_sz = '60' # 30pt
            elif max_u > 11:
                u_sz = '70' # 35pt
            else:
                u_sz = '76' # 38pt
                
        if force_nt_sz:
            nt_sz = str(force_nt_sz * 2)
        else:
            nt_len = len(full_name_title)
            if nt_len > 10:
                nt_sz = '80'  # 40pt
            elif nt_len > 7:
                nt_sz = '96'  # 48pt
            else:
                nt_sz = '108' # 54pt
                
        return u_sz, nt_sz

    def populate_card(self, tbl: ET.Element, unit_str: str, name_str: str, title_str: str,
                      custom_title: Optional[str] = None, custom_subtitle: Optional[str] = None,
                      force_u_sz: Optional[int] = None, force_nt_sz: Optional[int] = None):
        """Fills attendee details and optional custom headers into a card table."""
        u_lines = self.split_unit_text(unit_str)
        # Format name + title
        if name_str and title_str:
            full_name_title = f"{name_str} {title_str}" if len(name_str) <= 3 and len(title_str) <= 4 and ' ' not in name_str else f"{name_str}{title_str}"
        elif name_str:
            full_name_title = name_str
        else:
            full_name_title = title_str or ""

        u_sz, nt_sz = self.calculate_font_sizes(u_lines, full_name_title, force_u_sz, force_nt_sz)
        is_blank = not u_lines and not full_name_title

        # 1. Update DrawingML wsp textboxes
        for wsp in tbl.findall(f'.//{{{WPS_NS}}}wsp'):
            cNvPr = wsp.find(f'{{{WPS_NS}}}cNvPr')
            w_name = cNvPr.get('name', '') if cNvPr is not None else ''
            
            # Optional Header updates
            if custom_title and 'Rectangle 228' in w_name:
                for t in wsp.findall(f'.//{{{W_NS}}}t'):
                    t.text = custom_title
            if custom_subtitle and 'Rectangle 230' in w_name:
                t_nodes = wsp.findall(f'.//{{{W_NS}}}t')
                if t_nodes:
                    t_nodes[0].text = custom_subtitle
                    for extra_t in t_nodes[1:]:
                        extra_t.text = ""

            # Attendee Text Box
            if '文字方塊' in w_name:
                txbxContent = wsp.find(f'.//{{{W_NS}}}txbxContent')
                if txbxContent is not None:
                    txbxContent.clear()
                    if is_blank:
                        # Add empty paragraph with centered font
                        p_empty = ET.SubElement(txbxContent, f'{{{W_NS}}}p')
                        pPr = ET.SubElement(p_empty, f'{{{W_NS}}}pPr')
                        ET.SubElement(pPr, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'center'})
                        continue

                    # Unit Paragraph
                    p_unit = ET.SubElement(txbxContent, f'{{{W_NS}}}p')
                    pPr = ET.SubElement(p_unit, f'{{{W_NS}}}pPr')
                    ET.SubElement(pPr, f'{{{W_NS}}}spacing', {
                        f'{{{W_NS}}}beforeLines': '20',
                        f'{{{W_NS}}}before': '72',
                        f'{{{W_NS}}}line': '900',
                        f'{{{W_NS}}}lineRule': 'exact'
                    })
                    ET.SubElement(pPr, f'{{{W_NS}}}ind', {
                        f'{{{W_NS}}}leftChars': '100',
                        f'{{{W_NS}}}left': '240'
                    })
                    rPr = ET.SubElement(pPr, f'{{{W_NS}}}rPr')
                    ET.SubElement(rPr, f'{{{W_NS}}}rFonts', {
                        f'{{{W_NS}}}ascii': '標楷體',
                        f'{{{W_NS}}}eastAsia': '標楷體',
                        f'{{{W_NS}}}hAnsi': '標楷體'
                    })
                    ET.SubElement(rPr, f'{{{W_NS}}}b')
                    ET.SubElement(rPr, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': u_sz})
                    ET.SubElement(rPr, f'{{{W_NS}}}szCs', {f'{{{W_NS}}}val': u_sz})
                    glow = ET.SubElement(rPr, f'{{{W14_NS}}}glow', {f'{{{W14_NS}}}rad': '101600'})
                    sClr = ET.SubElement(glow, f'{{{W14_NS}}}schemeClr', {f'{{{W14_NS}}}val': 'accent3'})
                    ET.SubElement(sClr, f'{{{W14_NS}}}alpha', {f'{{{W14_NS}}}val': '60000'})
                    ET.SubElement(sClr, f'{{{W14_NS}}}satMod', {f'{{{W14_NS}}}val': '175000'})

                    for idx, line in enumerate(u_lines):
                        if idx > 0:
                            r_br = ET.SubElement(p_unit, f'{{{W_NS}}}r')
                            ET.SubElement(r_br, f'{{{W_NS}}}br')
                        r_t = ET.SubElement(p_unit, f'{{{W_NS}}}r')
                        r_t.append(copy.deepcopy(rPr))
                        t_node = ET.SubElement(r_t, f'{{{W_NS}}}t')
                        t_node.text = line

                    # Name + Title Paragraph
                    p_name = ET.SubElement(txbxContent, f'{{{W_NS}}}p')
                    pPr2 = ET.SubElement(p_name, f'{{{W_NS}}}pPr')
                    ET.SubElement(pPr2, f'{{{W_NS}}}widowControl')
                    ET.SubElement(pPr2, f'{{{W_NS}}}snapToGrid', {f'{{{W_NS}}}val': '0'})
                    ET.SubElement(pPr2, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'center'})
                    rPr2 = ET.SubElement(pPr2, f'{{{W_NS}}}rPr')
                    ET.SubElement(rPr2, f'{{{W_NS}}}rFonts', {
                        f'{{{W_NS}}}ascii': '標楷體',
                        f'{{{W_NS}}}eastAsia': '標楷體',
                        f'{{{W_NS}}}hAnsi': '標楷體'
                    })
                    ET.SubElement(rPr2, f'{{{W_NS}}}color', {
                        f'{{{W_NS}}}val': '000000',
                        f'{{{W_NS}}}themeColor': 'text1'
                    })
                    ET.SubElement(rPr2, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': nt_sz})
                    ET.SubElement(rPr2, f'{{{W_NS}}}szCs', {f'{{{W_NS}}}val': nt_sz})

                    r_nt = ET.SubElement(p_name, f'{{{W_NS}}}r')
                    r_nt.append(copy.deepcopy(rPr2))
                    t_nt = ET.SubElement(r_nt, f'{{{W_NS}}}t')
                    t_nt.text = full_name_title

        # 2. Update Fallback VML shapes
        for shape in tbl.findall(f'.//{{{V_NS}}}shape'):
            txbxContent = shape.find(f'.//{{{W_NS}}}txbxContent')
            if txbxContent is not None:
                txbxContent.clear()
                if is_blank:
                    continue
                # Unit p
                p_unit = ET.SubElement(txbxContent, f'{{{W_NS}}}p')
                pPr = ET.SubElement(p_unit, f'{{{W_NS}}}pPr')
                rPr = ET.SubElement(pPr, f'{{{W_NS}}}rPr')
                ET.SubElement(rPr, f'{{{W_NS}}}rFonts', {
                    f'{{{W_NS}}}ascii': '標楷體',
                    f'{{{W_NS}}}eastAsia': '標楷體',
                    f'{{{W_NS}}}hAnsi': '標楷體'
                })
                ET.SubElement(rPr, f'{{{W_NS}}}b')
                ET.SubElement(rPr, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': u_sz})
                for idx, line in enumerate(u_lines):
                    if idx > 0:
                        r_br = ET.SubElement(p_unit, f'{{{W_NS}}}r')
                        ET.SubElement(r_br, f'{{{W_NS}}}br')
                    r_t = ET.SubElement(p_unit, f'{{{W_NS}}}r')
                    t_node = ET.SubElement(r_t, f'{{{W_NS}}}t')
                    t_node.text = line

                # Name p
                p_name = ET.SubElement(txbxContent, f'{{{W_NS}}}p')
                pPr2 = ET.SubElement(p_name, f'{{{W_NS}}}pPr')
                ET.SubElement(pPr2, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'center'})
                rPr2 = ET.SubElement(pPr2, f'{{{W_NS}}}rPr')
                ET.SubElement(rPr2, f'{{{W_NS}}}rFonts', {
                    f'{{{W_NS}}}ascii': '標楷體',
                    f'{{{W_NS}}}eastAsia': '標楷體',
                    f'{{{W_NS}}}hAnsi': '標楷體'
                })
                ET.SubElement(rPr2, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': nt_sz})
                r_nt = ET.SubElement(p_name, f'{{{W_NS}}}r')
                t_nt = ET.SubElement(r_nt, f'{{{W_NS}}}t')
                t_nt.text = full_name_title

    def generate(self, attendees: List[Dict[str, str]], output_path: str,
                 blank_cards: int = 2, custom_title: Optional[str] = None,
                 custom_subtitle: Optional[str] = None, force_u_sz: Optional[int] = None,
                 force_nt_sz: Optional[int] = None) -> str:
        """Generates a complete meeting tent card document from a list of attendees."""
        out_p = Path(output_path).resolve()
        out_p.parent.mkdir(parents=True, exist_ok=True)

        body = self.root.find(f'{{{W_NS}}}body')
        body.clear()

        # 1. Populate attendee cards
        for att in attendees:
            card_tbl = copy.deepcopy(self.proto_tbl)
            self.populate_card(
                card_tbl,
                unit_str=att.get('unit', ''),
                name_str=att.get('name', ''),
                title_str=att.get('title', ''),
                custom_title=custom_title,
                custom_subtitle=custom_subtitle,
                force_u_sz=force_u_sz,
                force_nt_sz=force_nt_sz
            )
            body.append(card_tbl)
            body.append(copy.deepcopy(self.proto_p))

        # 2. Append blank padding cards
        for _ in range(blank_cards):
            blank_tbl = copy.deepcopy(self.proto_tbl)
            self.populate_card(
                blank_tbl,
                unit_str='',
                name_str='',
                title_str='',
                custom_title=custom_title,
                custom_subtitle=custom_subtitle
            )
            body.append(blank_tbl)
            body.append(copy.deepcopy(self.proto_p))

        # 3. Add document termination elements
        body.append(copy.deepcopy(self.final_p))
        body.append(copy.deepcopy(self.final_sectPr))

        # Write output docx
        new_doc_xml = ET.tostring(self.root, encoding='utf-8', xml_declaration=True)
        file_map = dict(self.file_map)
        file_map['word/document.xml'] = new_doc_xml

        with zipfile.ZipFile(out_p, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for name, content in file_map.items():
                z_out.writestr(name, content)

        return str(out_p)

    @classmethod
    def inspect(cls, docx_path: str) -> Dict[str, Any]:
        """Inspects an existing tent card docx file and extracts roster details."""
        doc_path = Path(docx_path).resolve()
        if not doc_path.exists():
            raise FileNotFoundError(f"File not found: {doc_path}")

        with zipfile.ZipFile(doc_path) as z:
            doc_xml = z.read('word/document.xml')

        root = ET.fromstring(doc_xml)
        tables = root.findall(f'.//{{{W_NS}}}tbl')

        roster = []
        blank_count = 0

        for idx, tbl in enumerate(tables):
            # Find attendee textbox
            wsp_matches = []
            for wsp in tbl.findall(f'.//{{{WPS_NS}}}wsp'):
                cNvPr = wsp.find(f'{{{WPS_NS}}}cNvPr')
                name = cNvPr.get('name', '') if cNvPr is not None else ''
                if '文字方塊' in name:
                    p_texts = []
                    for p in wsp.findall(f'.//{{{W_NS}}}p'):
                        t = ''.join([elem.text for elem in p.findall(f'.//{{{W_NS}}}t') if elem.text])
                        if t:
                            p_texts.append(t)
                    if p_texts:
                        wsp_matches.append(p_texts)
            
            if not wsp_matches:
                blank_count += 1
                roster.append({
                    "page": idx + 1,
                    "is_blank": True,
                    "unit": "",
                    "name_title": ""
                })
            else:
                card_data = wsp_matches[0]
                if len(card_data) == 1:
                    unit_val = ""
                    nt_val = card_data[0]
                else:
                    unit_val = " ".join(card_data[:-1])
                    nt_val = card_data[-1]

                roster.append({
                    "page": idx + 1,
                    "is_blank": False,
                    "unit": unit_val,
                    "name_title": nt_val
                })

        return {
            "file": str(doc_path),
            "total_cards": len(tables),
            "filled_cards": len(tables) - blank_count,
            "blank_cards": blank_count,
            "roster": roster
        }

    @classmethod
    def update_card(cls, source_docx: str, page: Optional[int] = None,
                    match_name: Optional[str] = None, unit: Optional[str] = None,
                    name: Optional[str] = None, title: Optional[str] = None,
                    output_docx: Optional[str] = None) -> str:
        """Updates a specific attendee card by page index or by matching name."""
        out_p = Path(output_docx or source_docx).resolve()

        with zipfile.ZipFile(source_docx) as z:
            file_map = {fname: z.read(fname) for fname in z.namelist()}

        root = ET.fromstring(file_map['word/document.xml'])
        tables = root.findall(f'.//{{{W_NS}}}tbl')

        target_idx = -1
        if page is not None:
            if 1 <= page <= len(tables):
                target_idx = page - 1
            else:
                raise IndexError(f"Page {page} out of range (1 to {len(tables)})")
        elif match_name:
            for idx, tbl in enumerate(tables):
                for wsp in tbl.findall(f'.//{{{WPS_NS}}}wsp'):
                    cNvPr = wsp.find(f'{{{WPS_NS}}}cNvPr')
                    if '文字方塊' in (cNvPr.get('name', '') if cNvPr is not None else ''):
                        txt = ''.join([t.text for t in wsp.findall(f'.//{{{W_NS}}}t') if t.text])
                        if match_name in txt:
                            target_idx = idx
                            break
                if target_idx != -1:
                    break
            if target_idx == -1:
                raise ValueError(f"No attendee matching '{match_name}' found in document.")
        else:
            raise ValueError("Either --page or --match-name must be specified to update.")

        # Inspect current table to get fallback values
        curr_tbl = tables[target_idx]
        current_unit = ""
        current_name_title = ""
        for wsp in curr_tbl.findall(f'.//{{{WPS_NS}}}wsp'):
            cNvPr = wsp.find(f'{{{WPS_NS}}}cNvPr')
            if '文字方塊' in (cNvPr.get('name', '') if cNvPr is not None else ''):
                p_texts = []
                for p in wsp.findall(f'.//{{{W_NS}}}p'):
                    t = ''.join([t_elem.text for t_elem in p.findall(f'.//{{{W_NS}}}t') if t_elem.text])
                    if t:
                        p_texts.append(t)
                if len(p_texts) >= 2:
                    current_unit = ' '.join(p_texts[:-1])
                    current_name_title = p_texts[-1]
                elif len(p_texts) == 1:
                    current_name_title = p_texts[0]
                break

        # Parse current_name and current_title from current_name_title
        curr_name = current_name_title
        curr_title = ""
        if " " in current_name_title:
            parts = current_name_title.split(" ", 1)
            curr_name, curr_title = parts[0], parts[1]
        else:
            known_titles = [
                '特聘講座教授', '講座教授', '特聘教授', '副教授', '助理教授', '教授',
                '課程督學', '專案督學', '督學', '候用校長', '校長', '教務主任', '學務主任',
                '總務主任', '主任', '輔導員', '專業助理', '助理研究員', '副研究員',
                '研究員', '科員', '專員', '老師', '組長', '秘書', '助理'
            ]
            for kt in known_titles:
                if current_name_title.endswith(kt) and len(current_name_title) > len(kt):
                    curr_name = current_name_title[:-len(kt)]
                    curr_title = kt
                    break

        final_u = unit if unit is not None else current_unit
        final_n = name if name is not None else curr_name
        final_t = title if title is not None else curr_title

        engine = cls(template_path=source_docx)
        engine.populate_card(
            curr_tbl,
            unit_str=final_u,
            name_str=final_n,
            title_str=final_t
        )

        new_doc_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        file_map['word/document.xml'] = new_doc_xml

        with zipfile.ZipFile(out_p, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for fname, content in file_map.items():
                z_out.writestr(fname, content)

        return str(out_p)

    @classmethod
    def append_attendees(cls, source_docx: str, new_attendees: List[Dict[str, str]], 
                         output_docx: Optional[str] = None) -> str:
        """Appends new attendees to an existing docx before the blank cards."""
        out_p = Path(output_docx or source_docx).resolve()
        
        with zipfile.ZipFile(source_docx) as z:
            file_map = {name: z.read(name) for name in z.namelist()}
            
        root = ET.fromstring(file_map['word/document.xml'])
        body = root.find(f'{{{W_NS}}}body')
        children = list(body)
        
        # Clone prototype table and section break p
        proto_tbl = copy.deepcopy(children[0])
        proto_p = copy.deepcopy(children[1])
        
        # Find index of first blank card, or append before final elements
        tables = root.findall(f'.//{{{W_NS}}}tbl')
        insert_idx = len(children) - 2 # Default before final p & sectPr
        
        for idx, tbl in enumerate(tables):
            # Check if blank
            has_text = False
            for wsp in tbl.findall(f'.//{{{WPS_NS}}}wsp'):
                cNvPr = wsp.find(f'{{{WPS_NS}}}cNvPr')
                if '文字方塊' in (cNvPr.get('name', '') if cNvPr is not None else ''):
                    for t in wsp.findall(f'.//{{{W_NS}}}t'):
                        if t.text and t.text.strip():
                            has_text = True
                            break
            if not has_text:
                # Find position of this table in body children
                try:
                    insert_idx = children.index(tbl)
                    break
                except ValueError:
                    pass

        engine = cls(template_path=source_docx)
        
        # Prepare new elements
        new_elements = []
        for att in new_attendees:
            new_tbl = copy.deepcopy(proto_tbl)
            engine.populate_card(
                new_tbl,
                unit_str=att.get('unit', ''),
                name_str=att.get('name', ''),
                title_str=att.get('title', '')
            )
            new_elements.append(new_tbl)
            new_elements.append(copy.deepcopy(proto_p))

        # Insert into body
        for elem in reversed(new_elements):
            body.insert(insert_idx, elem)

        # Write output
        new_doc_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        file_map['word/document.xml'] = new_doc_xml
        
        with zipfile.ZipFile(out_p, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for name, content in file_map.items():
                z_out.writestr(name, content)

        return str(out_p)

    @classmethod
    def parse_roster_text(cls, text: str) -> List[Dict[str, str]]:
        """Parses attendee list from raw string (newlines, literal \\n, or semicolons)."""
        attendees = []
        normalized = text.replace('\\n', '\n').replace(';', '\n')
        lines = [l.strip() for l in normalized.split('\n') if l.strip()]
        for line in lines:
            # Check delimiters
            if '|' in line:
                parts = [p.strip() for p in line.split('|') if p.strip()]
            elif '\t' in line:
                parts = [p.strip() for p in line.split('\t') if p.strip()]
            elif ',' in line:
                parts = [p.strip() for p in line.split(',') if p.strip()]
            elif ' / ' in line:
                parts = [p.strip() for p in line.split(' / ') if p.strip()]
            else:
                parts = [p.strip() for p in line.split() if p.strip()]

            if not parts:
                continue
            if len(parts) == 1:
                attendees.append({'unit': '', 'name': parts[0], 'title': ''})
            elif len(parts) == 2:
                attendees.append({'unit': parts[0], 'name': parts[1], 'title': ''})
            else:
                attendees.append({'unit': parts[0], 'name': parts[1], 'title': ' '.join(parts[2:])})
        return attendees

    @staticmethod
    def parse_input_file(file_path: str) -> List[Dict[str, str]]:
        """Parses attendee list from Excel, CSV, TSV, or JSON."""
        p = Path(file_path).resolve()
        if not p.exists():
            raise FileNotFoundError(f"Input file not found: {p}")
        
        ext = p.suffix.lower()
        attendees = []

        if ext in ['.xlsx', '.xls']:
            try:
                import pandas as pd
                df = pd.read_excel(p)
                # Normalize column headers
                col_map = {}
                for col in df.columns:
                    c_str = str(col).strip().lower()
                    if any(k in c_str for k in ['單位', '服務單位', '處室', '學校', 'unit', 'org', 'department']):
                        col_map['unit'] = col
                    elif any(k in c_str for k in ['姓名', '名字', '貴賓', 'name']):
                        col_map['name'] = col
                    elif any(k in c_str for k in ['職稱', '職務', '身分', 'title', 'role', 'position']):
                        col_map['title'] = col

                for _, row in df.iterrows():
                    u = str(row.get(col_map.get('unit', ''), '')).strip()
                    n = str(row.get(col_map.get('name', ''), '')).strip()
                    t = str(row.get(col_map.get('title', ''), '')).strip()
                    if u == 'nan': u = ''
                    if n == 'nan': n = ''
                    if t == 'nan': t = ''
                    if n or u or t:
                        attendees.append({'unit': u, 'name': n, 'title': t})
            except Exception as e:
                raise RuntimeError(f"Error reading Excel file: {e}")

        elif ext in ['.csv', '.tsv', '.txt']:
            import csv
            delim = '\t' if ext == '.tsv' else ','
            with open(p, 'r', encoding='utf-8-sig', errors='ignore') as f:
                sample = f.read(2048)
                f.seek(0)
                if ext == '.txt' and '\t' in sample:
                    delim = '\t'
                elif ext == '.txt' and ',' in sample:
                    delim = ','
                elif ext == '.txt':
                    delim = None # Line by line parsing
                
                if delim:
                    reader = csv.reader(f, delimiter=delim)
                    rows = list(reader)
                    if rows:
                        header = [c.strip().lower() for c in rows[0]]
                        u_idx, n_idx, t_idx = -1, -1, -1
                        for i, h in enumerate(header):
                            if any(k in h for k in ['單位', 'unit', 'org', '處室']): u_idx = i
                            elif any(k in h for k in ['姓名', 'name']): n_idx = i
                            elif any(k in h for k in ['職稱', 'title', 'role']): t_idx = i
                        
                        start_row = 1 if (u_idx != -1 or n_idx != -1) else 0
                        if u_idx == -1 and n_idx == -1:
                            u_idx, n_idx, t_idx = 0, 1, 2

                        for r in rows[start_row:]:
                            if not any(r): continue
                            u = r[u_idx].strip() if u_idx < len(r) and u_idx >= 0 else ''
                            n = r[n_idx].strip() if n_idx < len(r) and n_idx >= 0 else ''
                            t = r[t_idx].strip() if t_idx < len(r) and t_idx >= 0 else ''
                            if u or n or t:
                                attendees.append({'unit': u, 'name': n, 'title': t})
                else:
                    # Space or vertical bar delimited lines
                    for line in f:
                        line = line.strip()
                        if not line: continue
                        parts = [p.strip() for p in line.replace('|', ' ').split() if p.strip()]
                        if len(parts) == 1:
                            attendees.append({'unit': '', 'name': parts[0], 'title': ''})
                        elif len(parts) == 2:
                            attendees.append({'unit': parts[0], 'name': parts[1], 'title': ''})
                        elif len(parts) >= 3:
                            attendees.append({'unit': parts[0], 'name': parts[1], 'title': ' '.join(parts[2:])})

        elif ext == '.json':
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    for item in data:
                        if isinstance(item, dict):
                            attendees.append({
                                'unit': str(item.get('unit', '')).strip(),
                                'name': str(item.get('name', '')).strip(),
                                'title': str(item.get('title', '')).strip()
                            })

        return attendees


def main():
    parser = argparse.ArgumentParser(description="Meeting Tent Card Engine (會議桌牌核心引擎)")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # generate command
    p_gen = subparsers.add_parser("generate", help="Generate tent cards docx from roster")
    p_gen.add_argument("-i", "--input", help="Input roster file (.xlsx, .csv, .txt, .json)")
    p_gen.add_argument("-r", "--roster", help="Raw roster string (e.g. '單位 | 姓名 | 職稱\\n...')")
    p_gen.add_argument("-o", "--output", required=True, help="Output docx file path")
    p_gen.add_argument("-t", "--template", help="Custom template docx path")
    p_gen.add_argument("-b", "--blanks", type=int, default=2, help="Number of blank cards to append (default: 2)")
    p_gen.add_argument("--event-title", help="Custom event title header")
    p_gen.add_argument("--sub-title", help="Custom organization subtitle")
    p_gen.add_argument("--unit-font-size", type=int, help="Force unit font size in pt (e.g. 38, 35)")
    p_gen.add_argument("--name-font-size", type=int, help="Force name font size in pt (e.g. 54, 48)")

    # inspect command
    p_ins = subparsers.add_parser("inspect", help="Inspect an existing tent card docx")
    p_ins.add_argument("-f", "--file", required=True, help="Target docx file path")
    p_ins.add_argument("--json", action="store_true", help="Output results as JSON")

    # append command
    p_app = subparsers.add_parser("append", help="Append new attendee(s) to existing docx")
    p_app.add_argument("-f", "--file", required=True, help="Target docx file path")
    p_app.add_argument("-o", "--output", help="Output file path (default: overwrite target)")
    p_app.add_argument("-u", "--unit", default="", help="Attendee affiliation/unit")
    p_app.add_argument("-n", "--name", default="", help="Attendee name")
    p_app.add_argument("-t", "--title", default="", help="Attendee title")
    p_app.add_argument("-i", "--input", help="File with additional attendees to append")

    # update command
    p_upd = subparsers.add_parser("update", help="Update attendee info on a specific card")
    p_upd.add_argument("-f", "--file", required=True, help="Target docx file path")
    p_upd.add_argument("-o", "--output", help="Output file path (default: overwrite target)")
    p_upd.add_argument("-p", "--page", type=int, help="1-based page number to update")
    p_upd.add_argument("-m", "--match-name", help="Name string to search and replace")
    p_upd.add_argument("-u", "--unit", help="Updated affiliation/unit")
    p_upd.add_argument("-n", "--name", help="Updated attendee name")
    p_upd.add_argument("-t", "--title", help="Updated attendee title")

    # export command
    p_exp = subparsers.add_parser("export", help="Export attendee roster from docx")
    p_exp.add_argument("-f", "--file", required=True, help="Target docx file path")
    p_exp.add_argument("-o", "--output", required=True, help="Output file (.csv, .xlsx, .json)")

    args = parser.parse_args()

    if args.command == "generate":
        engine = TentCardEngine(template_path=args.template)
        attendees = []
        if args.input:
            attendees = engine.parse_input_file(args.input)
        elif args.roster:
            attendees = engine.parse_roster_text(args.roster)
        if not attendees:
            print("Error: No attendee data provided (use -i or -r).", file=sys.stderr)
            sys.exit(1)
        
        out = engine.generate(
            attendees=attendees,
            output_path=args.output,
            blank_cards=args.blanks,
            custom_title=args.event_title,
            custom_subtitle=args.sub_title,
            force_u_sz=args.unit_font_size,
            force_nt_sz=args.name_font_size
        )
        print(f"Successfully generated tent cards ({len(attendees)} attendees + {args.blanks} blanks): {out}")

    elif args.command == "inspect":
        info = TentCardEngine.inspect(args.file)
        if args.json:
            print(json.dumps(info, ensure_ascii=False, indent=2))
        else:
            print(f"=== Meeting Tent Card Inspection ===")
            print(f"File: {info['file']}")
            print(f"Total Cards: {info['total_cards']} (Filled: {info['filled_cards']}, Blank: {info['blank_cards']})")
            print("-" * 50)
            for r in info['roster']:
                status = "[BLANK]" if r['is_blank'] else f"{r['unit']} | {r['name_title']}"
                print(f"Page {r['page']:02d}: {status}")

    elif args.command == "append":
        new_atts = []
        if args.input:
            new_atts = TentCardEngine.parse_input_file(args.input)
        elif args.name or args.unit or args.title:
            new_atts = [{'unit': args.unit, 'name': args.name, 'title': args.title}]
        
        if not new_atts:
            print("Error: No new attendee provided to append.", file=sys.stderr)
            sys.exit(1)

        out = TentCardEngine.append_attendees(
            source_docx=args.file,
            new_attendees=new_atts,
            output_docx=args.output
        )
        print(f"Successfully appended {len(new_atts)} attendees to: {out}")

    elif args.command == "update":
        out = TentCardEngine.update_card(
            source_docx=args.file,
            page=args.page,
            match_name=args.match_name,
            unit=args.unit,
            name=args.name,
            title=args.title,
            output_docx=args.output
        )
        print(f"Successfully updated card in: {out}")

    elif args.command == "export":
        info = TentCardEngine.inspect(args.file)
        out_p = Path(args.output).resolve()
        ext = out_p.suffix.lower()
        
        roster_data = [r for r in info['roster'] if not r['is_blank']]
        
        if ext == '.json':
            with open(out_p, 'w', encoding='utf-8') as f:
                json.dump(roster_data, f, ensure_ascii=False, indent=2)
        elif ext == '.csv':
            import csv
            with open(out_p, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Page', 'Unit', 'Name_Title'])
                for r in roster_data:
                    writer.writerow([r['page'], r['unit'], r['name_title']])
        elif ext in ['.xlsx', '.xls']:
            import pandas as pd
            df = pd.DataFrame(roster_data)[['page', 'unit', 'name_title']]
            df.columns = ['頁碼', '服務單位', '姓名職稱']
            df.to_excel(out_p, index=False)
        print(f"Exported {len(roster_data)} attendees to {out_p}")

    else:
        parser.print_help()


if __name__ == '__main__':
    main()
