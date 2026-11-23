#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
混铁车中夜班安全消防检查情况报告自动生成工具 - 核心逻辑 v1.0.0
"""

import logging
import os
import re
import sys
import zipfile
from datetime import date, datetime
from io import BytesIO
from xml.etree import ElementTree as ET

import openpyxl
import xlrd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt
from openpyxl.utils.datetime import from_excel


logger = logging.getLogger(__name__)


# ==================== 配置常量 ====================
# 事业部名称：写入报告标题与输出文件名
BI_NAME = '混铁车事业部'

# 隐患登记表未标注作业区时，列表显示的默认作业区
DEFAULT_AREA_LABEL = '鱼雷罐'

# 作业区关键字 -> 值班表键名（按顺序匹配，命中即返回）
AREA_KEY_MAPPING = (
    ('煤库', 'coal'),
    ('原料', 'raw_material'),
    ('供矿', 'ore'),
)

# 作业区键名 -> 中文标签
AREA_LABEL_MAPPING = {
    'coal': '煤库作业区',
    'raw_material': '原料分厂作业区',
    'ore': '供矿作业区',
}

# 值班领导姓名 -> 职务（未匹配到时默认“管理人员”）
LEADER_ROLE_MAPPING = {
    '王宝琪': '厂长',
    '蒋军': '副厂长',
    '冯共华': '厂长助理',
}
DEFAULT_ROLE = '管理人员'

# 值班表表头扫描行数
DUTY_HEADER_SCAN_ROWS = 12

# 隐患登记表表头扫描行数
SWEEP_HEADER_SCAN_ROWS = 15

# Word 模板填充坐标 (行, 列)
DOC_CELL_TITLE = 0                # 标题段落索引
DOC_CELL_LEADER = (2, 0)          # 值班领导姓名
DOC_CELL_LEADER_ROLE = (2, 3)     # 值班领导职务
DOC_CELL_INSPECTOR = (2, 5)       # 检查人姓名
DOC_CELL_INSPECTOR_ROLE = (2, 12) # 检查人职务
DOC_CELL_DATE = (12, 0)           # 检查日期
DOC_CELL_NATURE = (12, 1)         # 性质（隐患/违章）
DOC_CELL_DESC = (12, 2)           # 问题描述
DOC_CELL_NATURE2 = (12, 6)        # 性质（重复列）
DOC_CELL_IMAGE = (12, 12)         # 图片单元格
DOC_CELL_RESPONSIBLE = (12, 16)   # 责任人


# ==================== 通用小工具 ====================
def _cell_text(value):
    """把单元格值转成去空白的字符串（None -> ''）。"""
    if value is None:
        return ''
    return str(value).strip()


def _parse_date_value(value, datemode=None):
    """尽量把日期单元格（datetime/Excel序列号/字符串）解析为 datetime。"""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, (int, float)) and value:
        try:
            if datemode is not None:
                return datetime(*xlrd.xldate_as_tuple(value, datemode)[:6])
            return from_excel(value)
        except Exception:
            return None
    text = _cell_text(value)
    if not text:
        return None
    for fmt in ("%Y.%m.%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d %H:%M",
                "%Y-%m-%d %H:%M:%S", "%Y年%m月%d日"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _clean_name(value):
    """清理值班表中的姓名：去掉电话、括号、空白及残留的岗位标签。"""
    name = _cell_text(value)
    name = re.sub(r'[\d\s（）()\-]+', '', name)
    for label in ('煤库作业区', '原料分厂作业区', '供矿作业区',
                  '作业区值班员', '值班员', '值班领导'):
        name = name.replace(label, '')
    return name.strip()


def _read_sheet_rows(filepath):
    """读取 Excel 第一个工作表，返回 (行数据列表, datemode/epoch)。"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.xls':
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        rows = [[ws.cell_value(r, c) for c in range(ws.ncols)]
                for r in range(ws.nrows)]
        return rows, wb.datemode
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    try:
        ws = wb.active
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
        return rows, None
    finally:
        wb.close()


def _index_of(header, *keywords):
    """在表头行中查找包含任一关键字的列索引（0-based），找不到返回 None。"""
    for i, text in enumerate(header):
        if any(k in text for k in keywords):
            return i
    return None


# ==================== 值班表读取 ====================
def read_duty_roster(filepath):
    """读取值班表（.xls/.xlsx），返回日期->信息的字典。

    支持两种列结构：
    1) 混铁车值班表：时间 | 星期 | 值班领导 | 作业区值班员
    2) 炼铁值班表：日期 | 星期 | 值班领导 | 原料分厂 | 煤库作业区 | 供矿作业区

    返回字典同时包含完整日期（YYYY.M.D）和短日期（M.D）两种键。
    """
    rows, datemode = _read_sheet_rows(filepath)
    header_idx, header = None, None
    for idx, row in enumerate(rows[:DUTY_HEADER_SCAN_ROWS]):
        texts = [_cell_text(v) for v in row]
        joined = ' '.join(texts)
        if '值班领导' in joined and ('时间' in joined or '日期' in joined):
            header_idx, header = idx, texts
            break
    if header_idx is None:
        raise ValueError('未能识别值班表表头（需包含“时间/日期”和“值班领导”列）')

    col_date = _index_of(header, '时间', '日期')
    col_week = _index_of(header, '星期')
    col_leader = _index_of(header, '值班领导')
    col_inspector = _index_of(header, '作业区值班员', '值班员')
    col_raw = _index_of(header, '原料')
    col_coal = _index_of(header, '煤库')
    col_ore = _index_of(header, '供矿')

    def value_of(row, col):
        if col is None or col >= len(row):
            return ''
        return row[col]

    roster = {}
    for row_idx in range(header_idx + 1, len(rows)):
        row = rows[row_idx]
        raw_date = value_of(row, col_date)
        if raw_date is None or not str(raw_date).strip():
            continue
        dt = _parse_date_value(raw_date, datemode)
        if dt is None:
            logger.warning('值班表第%d行日期无法解析，已跳过: %r', row_idx + 1, raw_date)
            continue

        leader_name = _clean_name(value_of(row, col_leader))
        raw_name = _clean_name(value_of(row, col_raw))
        coal_name = _clean_name(value_of(row, col_coal))
        ore_name = _clean_name(value_of(row, col_ore))
        inspector_name = _clean_name(value_of(row, col_inspector))
        if not inspector_name:
            inspector_name = coal_name or raw_name or ore_name

        entry = {
            'leader': leader_name,
            'inspector': inspector_name,
            'raw_material': raw_name,
            'coal': coal_name,
            'ore': ore_name,
            'day_of_week': _cell_text(value_of(row, col_week)),
            'full_date': dt,
        }
        roster[dt.strftime('%Y.%m.%d')] = entry
        roster[f'{dt.month}.{dt.day}'] = entry
    return roster


def _normalize_md(month, day):
    """将月份/日期规范化为 'm.d'（去掉前导零）形式，保证键一致。"""
    try:
        return f"{int(month)}.{int(day)}"
    except (ValueError, TypeError):
        return f"{month}.{day}"


def resolve_duty(roster_data, month, day, area_key=None):
    """根据月份、日期查找值班信息，返回 (leader, inspector) 元组。

    area_key 指定时 inspector 取该作业区人员；未指定或未命中时取第一个非空人员。
    """
    duty = roster_data.get(_normalize_md(month, day), {})
    if not duty:
        zm = str(month).zfill(2)
        zd = str(day).zfill(2)
        duty = roster_data.get(f"{zm}.{zd}", {})
    if not duty:
        full_key = f"{datetime.now().year}.{str(month).zfill(2)}.{str(day).zfill(2)}"
        duty = roster_data.get(full_key, {})
    leader = duty.get('leader', '')
    inspector = duty.get(area_key, '') if area_key else ''
    if not inspector:
        for k in ('inspector', 'coal', 'raw_material', 'ore'):
            if duty.get(k):
                inspector = duty[k]
                break
    return leader, inspector



# ==================== 隐患登记表读取 ====================
def read_minesweeper(filepath):
    """读取隐患登记表（原“扫雷表”），自动识别格式，返回 (问题列表, 格式名称)。

    支持的格式：
    - 混铁车格式：表头含“事故隐患排查时间 / 隐患项目内容 / 整改前 / 整改后”
    - 格式2：标题含“排查治理”，数据从第6行开始
    - 格式1：原扫雷表，数据从第10行开始
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    try:
        ws = wb.active
        header_row, header = _find_sweep_header(ws)
        if header_row is not None:
            items = _read_huitieche_items(ws, header_row, header)
            format_name = '混铁车（事故隐患排查整改登记表）'
        else:
            title = _cell_text(ws.cell(1, 1).value)
            if '排查治理' in title:
                items = _read_legacy_format2(ws)
                format_name = '格式2（排查治理）'
            else:
                items = _read_legacy_format1(ws)
                format_name = '格式1（隐患排查）'
        _attach_images(wb, filepath, items)
        return items, format_name
    finally:
        wb.close()


def _find_sweep_header(ws):
    """查找混铁车隐患登记表的表头行，返回 (1-based行号, 表头文本列表)。"""
    max_row = min(ws.max_row, SWEEP_HEADER_SCAN_ROWS)
    for r in range(1, max_row + 1):
        texts = [_cell_text(ws.cell(r, c).value)
                 for c in range(1, ws.max_column + 1)]
        joined = ''.join(texts)
        if ('隐患项目内容' in joined
                and ('隐患排查时间' in joined or '检查时间' in joined)):
            return r, texts
    return None, None


def _read_huitieche_items(ws, header_row, header):
    """读取混铁车隐患登记表的数据行。"""
    col = {}
    for i, text in enumerate(header, start=1):
        if '序号' in text:
            col.setdefault('seq', i)
        elif '排查时间' in text or '检查时间' in text:
            col.setdefault('sweep_date', i)
        elif '发现人' in text:
            col.setdefault('finder', i)
        elif '地点' in text:
            col.setdefault('location', i)
        elif '项目内容' in text or '问题描述' in text:
            col.setdefault('description', i)
        elif '整改要求' in text or '整改措施' in text:
            col.setdefault('measure', i)
        elif '整改部门' in text:
            col.setdefault('responsible', i)
        elif '整改结果' in text:
            col.setdefault('completion', i)
        elif '验收人' in text:
            col.setdefault('verifier', i)
        elif '整改前' in text:
            col.setdefault('before', i)
        elif '整改后' in text:
            col.setdefault('after', i)
    if 'description' not in col:
        raise ValueError('隐患登记表缺少“隐患项目内容”列')

    def value(row, key):
        c = col.get(key)
        if c is None:
            return ''
        return ws.cell(row, c).value

    items = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        description = _cell_text(value(row_idx, 'description'))
        if not description:
            continue  # 只有序号、没有内容的空行不计入
        seq_val = value(row_idx, 'seq')
        try:
            seq = int(seq_val)
        except (TypeError, ValueError):
            seq = row_idx - header_row
        dt = _parse_date_value(value(row_idx, 'sweep_date'))
        sweep_date = (f'{dt.year}.{dt.month}.{dt.day}' if dt
                      else _cell_text(value(row_idx, 'sweep_date')))
        location = _cell_text(value(row_idx, 'location'))
        finder = _cell_text(value(row_idx, 'finder'))
        responsible = _cell_text(value(row_idx, 'responsible')) or finder
        items.append({
            'seq': seq,
            'sweep_date': sweep_date,
            'contract_name': '',
            'supplier': '',
            'area': location,
            'location': location,
            'description': description,
            'finder': finder,
            'nature': '隐患',
            'measure': _cell_text(value(row_idx, 'measure')),
            'responsible': responsible,
            'verifier': _cell_text(value(row_idx, 'verifier')),
            'plan_date': '',
            'completion': _cell_text(value(row_idx, 'completion')),
            'remaining_issue': '',
            '_row': row_idx,
            '_before_col': col.get('before'),
            '_after_col': col.get('after'),
        })
    return items



def _read_legacy_format2(ws):
    """读取格式2（“扫雷”排查治理情况统计表），数据从第6行开始。"""
    items = []
    for row_idx in range(6, ws.max_row + 1):
        seq = ws.cell(row_idx, 1).value
        if seq is None:
            continue
        try:
            seq = int(seq)
        except (ValueError, TypeError):
            continue

        date_val = ws.cell(row_idx, 2).value
        sweep_date = _cell_text(date_val)
        dt = _parse_date_value(date_val)
        if dt:
            sweep_date = f'{dt.year}.{dt.month}.{dt.day}'

        cat_mgmt = _cell_text(ws.cell(row_idx, 11).value)
        cat_state = _cell_text(ws.cell(row_idx, 12).value)
        cat_env = _cell_text(ws.cell(row_idx, 13).value)
        nature = '一般'
        if cat_mgmt in ('✓', 'þ', '√'):
            nature = '管理缺陷'
        elif cat_state in ('✓', 'þ', '√'):
            nature = '物不安全状态'
        elif cat_env in ('✓', 'þ', '√'):
            nature = '环境因素'

        measure = _cell_text(ws.cell(row_idx, 14).value)
        plan = _cell_text(ws.cell(row_idx, 16).value)
        done = _cell_text(ws.cell(row_idx, 17).value)
        full_measure = measure or plan or done

        items.append({
            'seq': seq,
            'sweep_date': sweep_date,
            'contract_name': _cell_text(ws.cell(row_idx, 3).value),
            'supplier': '',
            'area': _cell_text(ws.cell(row_idx, 4).value),
            'location': _cell_text(ws.cell(row_idx, 5).value),
            'description': _cell_text(ws.cell(row_idx, 6).value),
            'finder': _cell_text(ws.cell(row_idx, 9).value),
            'nature': nature,
            'category_mgmt': cat_mgmt,
            'category_unsafe_state': cat_state,
            'category_unsafe_behavior': cat_env,
            'measure': full_measure,
            'responsible': _cell_text(ws.cell(row_idx, 9).value),
            'plan_date': done,
            'completion': done,
            'remaining_issue': '',
            '_row': row_idx,
        })
    return items


def _read_legacy_format1(ws):
    """读取格式1（原扫雷表），数据从第10行开始。"""
    items = []
    for row_idx in range(10, ws.max_row + 1):
        seq = ws.cell(row_idx, 1).value
        if seq is None:
            continue
        try:
            seq = int(seq)
        except (ValueError, TypeError):
            continue
        items.append({
            'seq': seq,
            'sweep_date': '',
            'contract_name': _cell_text(ws.cell(row_idx, 2).value),
            'supplier': _cell_text(ws.cell(row_idx, 3).value),
            'area': _cell_text(ws.cell(row_idx, 4).value),
            'location': _cell_text(ws.cell(row_idx, 5).value),
            'description': _cell_text(ws.cell(row_idx, 6).value),
            'finder': _cell_text(ws.cell(row_idx, 9).value),
            'nature': _cell_text(ws.cell(row_idx, 10).value),
            'category_mgmt': _cell_text(ws.cell(row_idx, 11).value),
            'category_unsafe_state': _cell_text(ws.cell(row_idx, 12).value),
            'category_unsafe_behavior': _cell_text(ws.cell(row_idx, 13).value),
            'measure': _cell_text(ws.cell(row_idx, 14).value),
            'responsible': _cell_text(ws.cell(row_idx, 15).value),
            'plan_date': _cell_text(ws.cell(row_idx, 16).value),
            'completion': _cell_text(ws.cell(row_idx, 17).value),
            'remaining_issue': _cell_text(ws.cell(row_idx, 18).value),
            '_row': row_idx,
        })
    return items



# ==================== 图片提取 ====================
def _collect_row_images(wb, filepath):
    """提取工作表中的图片，返回 {0-based行号: [{'col': 0-based列号, ...}]}。

    单元格图片（WPS DISPIMG）只记录压缩包内路径，生成报告时再按需读取，
    避免一次性把上百兆图片全部载入内存。
    """
    ws = wb.active
    row_images = {}

    # 1) 传统浮动图片（openpyxl 已在内存中）
    if hasattr(ws, '_images') and ws._images:
        for img in ws._images:
            try:
                r, c = img.anchor._from.row, img.anchor._from.col
                row_images.setdefault(r, []).append({'col': c, 'data': img._data()})
            except Exception:
                logger.warning('传统浮动图片提取失败', exc_info=True)

    # 2) “放入单元格”图片 (cellimages.xml)
    if filepath and os.path.exists(filepath):
        try:
            with zipfile.ZipFile(filepath, 'r') as z:
                if 'xl/cellimages.xml' not in z.namelist():
                    return row_images

                # 预扫描：DISPIMG ID -> (row, col)
                id_to_cell = {}
                for ri in range(1, ws.max_row + 1):
                    for ci in range(1, ws.max_column + 1):
                        v = _cell_text(ws.cell(ri, ci).value)
                        if 'DISPIMG' in v:
                            m = re.search(r'DISPIMG\("([^"]+)"', v)
                            if m:
                                id_to_cell[m.group(1)] = (ri - 1, ci - 1)

                if not id_to_cell:
                    return row_images

                # 解析 rels: rId -> 压缩包内路径
                id2file = {}
                for rel in ET.fromstring(z.read('xl/_rels/cellimages.xml.rels')):
                    rid, tgt = rel.get('Id'), rel.get('Target', '')
                    if rid and tgt:
                        id2file[rid] = 'xl/' + tgt.replace('..', '').lstrip('/')

                ns = {
                    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                    'etc': 'http://www.wps.cn/officeDocument/2017/etCustomData',
                }
                r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

                for ci in ET.fromstring(z.read('xl/cellimages.xml')).findall('etc:cellImage', ns):
                    nv = ci.find('.//xdr:cNvPr', ns)
                    blip = ci.find('.//a:blip', ns)
                    if nv is None or blip is None:
                        continue
                    img_id = nv.get('name', '')
                    rid = blip.get('{%s}embed' % r_ns, '')
                    img_file = id2file.get(rid, '')
                    cell = id_to_cell.get(img_id)
                    if img_id and img_file and cell:
                        row_images.setdefault(cell[0], []).append({
                            'col': cell[1],
                            'source': filepath,
                            'zip_path': img_file,
                        })
        except Exception:
            logger.warning('解析 cellimages.xml 失败', exc_info=True)
    return row_images


def _attach_images(wb, filepath, items):
    """把图片按实际行号关联到每个隐患项。"""
    row_images = _collect_row_images(wb, filepath)
    for item in items:
        row0 = item.get('_row', 0) - 1
        imgs = row_images.get(row0, [])
        if not imgs:
            continue
        before_col = item.get('_before_col')
        after_col = item.get('_after_col')
        if before_col or after_col:
            for im in imgs:
                if before_col and im['col'] == before_col - 1:
                    item.setdefault('images_before', []).append(im)
                elif after_col and im['col'] == after_col - 1:
                    item.setdefault('images_after', []).append(im)
                else:
                    item.setdefault('images', []).append(im)
        else:
            item['images'] = imgs


def _image_bytes(image):
    """读取图片二进制内容（浮动图片直接取 data，单元格图片按需从 xlsx 读取）。"""
    if image.get('data') is not None:
        return image['data']
    source, zip_path = image.get('source'), image.get('zip_path')
    if source and zip_path:
        with zipfile.ZipFile(source, 'r') as z:
            return z.read(zip_path)
    return None



# ==================== 作业区识别 ====================
def get_sweep_area_key(sweep_item):
    """根据隐患项的supplier/area字段，判断属于哪个作业区。

    返回: 'coal' | 'raw_material' | 'ore'；无法识别时返回 None（由调用方兜底）。
    """
    supplier = _cell_text(sweep_item.get('supplier', ''))
    area = _cell_text(sweep_item.get('area', ''))
    combined = supplier + ' ' + area

    for keyword, key in AREA_KEY_MAPPING:
        if keyword in combined:
            return key
    return None


def get_sweep_area_label(sweep_item):
    """获取作业区的中文标签。"""
    return AREA_LABEL_MAPPING.get(get_sweep_area_key(sweep_item), DEFAULT_AREA_LABEL)


# ==================== Word 报告生成 ====================
def _find_template():
    """定位模板文件路径，找不到时抛出 FileNotFoundError。"""
    base = os.path.dirname(os.path.abspath(__file__))
    meipass = getattr(sys, '_MEIPASS', None)  # PyInstaller 打包后路径
    candidates = ['模板.docx',
                  os.path.join(base, '模板.docx'),
                  os.path.join(base, '..', '模板.docx')]
    if meipass:
        candidates.insert(0, os.path.join(meipass, '模板.docx'))
    for c in candidates:
        if os.path.exists(c):
            return c
    raise FileNotFoundError('找不到模板.docx')


def _leader_role(leader_name):
    """根据值班领导姓名返回职务，未匹配到时返回默认职务。"""
    name = str(leader_name or '').strip()
    if name in LEADER_ROLE_MAPPING:
        return LEADER_ROLE_MAPPING[name]
    for key, role in LEADER_ROLE_MAPPING.items():
        if key and key in name:
            return role
    return DEFAULT_ROLE


def generate_docx(roster_data, sweep_item, output_path, month, day_str, template_path=None):
    month = str(month).strip()
    day_str = str(day_str).strip()
    day_num = day_str.split('.')[1] if '.' in day_str else day_str
    if template_path is None:
        template_path = _find_template()
    doc = Document(template_path)
    table = doc.tables[0]
    title_text = (f'    {BI_NAME}   {month}月{day_num}日中夜班'
                  f'（双休日、节假日）安全消防检查情况')
    if doc.paragraphs and doc.paragraphs[0].runs:
        for run in doc.paragraphs[0].runs:
            run.text = ''
        doc.paragraphs[0].runs[0].text = title_text
    area_key = get_sweep_area_key(sweep_item)
    leader_name, inspector_name = resolve_duty(roster_data, month, day_num, area_key)
    _fill_cell_text(table.cell(*DOC_CELL_LEADER), leader_name)
    _fill_cell_text(table.cell(*DOC_CELL_LEADER_ROLE), _leader_role(leader_name))
    _fill_cell_text(table.cell(*DOC_CELL_INSPECTOR), inspector_name)
    _fill_cell_text(table.cell(*DOC_CELL_INSPECTOR_ROLE), DEFAULT_ROLE)
    _fill_cell_text(table.cell(*DOC_CELL_DATE), f'{month}.{day_num}')
    nat = '违章' if sweep_item.get('nature', '') and '违章' in sweep_item['nature'] else '隐患'
    _fill_cell_text(table.cell(*DOC_CELL_NATURE), nat)
    _fill_cell_text(table.cell(*DOC_CELL_DESC), sweep_item.get('description', ''))
    _fill_cell_text(table.cell(*DOC_CELL_NATURE2), nat)
    _fill_cell_text(table.cell(*DOC_CELL_RESPONSIBLE), '责任人：' + sweep_item.get('responsible', ''))
    img_cell = table.cell(*DOC_CELL_IMAGE)
    for p in list(img_cell.paragraphs):
        p._element.getparent().remove(p._element)
    img_cell._tc.append(parse_xml(f"<w:p {nsdecls('w')}></w:p>"))
    before_imgs, after_imgs = _split_images(sweep_item)
    for idx, im in enumerate(before_imgs + after_imgs):
        try:
            if idx == 0 and before_imgs:
                p = img_cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run('整改前')
                r.font.size = Pt(9)
            if idx == len(before_imgs) and after_imgs:
                p = img_cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run('整改后')
                r.font.size = Pt(9)
            data = _image_bytes(im)
            if not data:
                continue
            p = img_cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run()
            r.add_picture(BytesIO(data), width=Inches(1.5))
        except Exception:
            logger.warning('第%s项图片写入失败', sweep_item.get('seq'), exc_info=True)
    doc.save(output_path)



def _split_images(sweep_item):
    """区分整改前/整改后图片，兼容旧格式按列位置判断的方式。"""
    before = sweep_item.get('images_before')
    after = sweep_item.get('images_after')
    if before is not None or after is not None:
        return list(before or []), list(after or [])
    images = sweep_item.get('images', [])
    before = [i for i in images if i.get('col', 0) <= 8]
    after = [i for i in images if i.get('col', 0) >= 15]
    return before, after


def _fill_cell_text(cell, text):
    """写入文本，保留模板单元格原有 run 的字体/字号格式（仿宋·小三）。

    做法：保留第一个 run（含其 rPr 格式），删除其余多余的空 run，再改文字，
    避免新建 run 时丢失模板的字体、字号等格式。
    """
    p = cell.paragraphs[0]
    runs = p.runs
    if runs:
        first = runs[0]
        for run in list(runs[1:]):
            run._element.getparent().remove(run._element)
        first.text = str(text)
    else:
        p.add_run(str(text))

