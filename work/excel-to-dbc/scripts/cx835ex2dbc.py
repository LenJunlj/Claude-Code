#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cx835ex2dbc.py — CX835EX VP LDCANFD FSCM 通信矩阵 → CANoe DBC 转换器

直接从 xlsx ZIP 中解析 XML (无需 openpyxl 依赖)，生成 ANSI 编码的 DBC 文件。
"""

import sys, os, re, zipfile, xml.etree.ElementTree as ET

# ════════════════════════════════════════════════════════════════
# XLSX 解析 (纯标准库, 无 openpyxl 依赖)
# ════════════════════════════════════════════════════════════════

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def parse_xlsx(filepath):
    """读取 xlsx 返回 {(row, col): value} 字典"""
    z = zipfile.ZipFile(filepath)
    # 共享字符串表
    ss_tree = ET.parse(z.open('xl/sharedStrings.xml'))
    ss_root = ss_tree.getroot()
    shared_strings = []
    for si in ss_root.findall('.//s:si', NS):
        texts = si.findall('.//s:t', NS)
        shared_strings.append(''.join(t.text or '' for t in texts))
    # 确定 FSCM_LDCANFD 对应的工作表文件
    wb_tree = ET.parse(z.open('xl/workbook.xml'))
    wb_root = wb_tree.getroot()
    sheet_target = None
    for sheet in wb_root.findall('.//s:sheet', NS):
        if sheet.get('name') == 'FSCM_LDCANFD':
            sheet_target = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            break
    if not sheet_target:
        # 回退到第一个 sheet
        sheet_target = 'rId1'
    # 找对应的 worksheet 文件
    rels_tree = ET.parse(z.open('xl/_rels/workbook.xml.rels'))
    rels_root = rels_tree.getroot()
    for rel in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        if rel.get('Id') == sheet_target:
            sheet_file = 'xl/' + rel.get('Target')
            break
    else:
        # 直接尝试 sheet4
        sheet_file = 'xl/worksheets/sheet4.xml'
    # 解析工作表
    data = {}
    sheet_tree = ET.parse(z.open(sheet_file))
    sheet_root = sheet_tree.getroot()
    for row in sheet_root.findall('.//s:row', NS):
        r = int(row.get('r'))
        for c in row.findall('.//s:c', NS):
            ref = c.get('r')
            col_str = ''.join(ch for ch in ref if ch.isalpha())
            t = c.get('t', '')
            v = c.find('s:v', NS)
            if v is not None:
                if t == 's':
                    idx = int(v.text)
                    val = shared_strings[idx] if idx < len(shared_strings) else ''
                else:
                    val = v.text or ''
            else:
                val = ''
            data[(r, col_str)] = val
    z.close()
    return data


# ════════════════════════════════════════════════════════════════
# 辅助函数
# ════════════════════════════════════════════════════════════════

def get_cell(data, row, col):
    """安全获取单元格值，去除_x000D_和多余空白，不保留换行"""
    v = data.get((row, col), '').strip()
    v = v.replace('_x000D_', '').replace('\r', '').strip()
    return v


def get_cell_raw(data, row, col):
    """获取原始单元格值，保留换行 (用于 VAL 描述等)"""
    v = data.get((row, col), '').strip()
    v = v.replace('_x000D_', '')  # 只去掉回车，保留 \n
    return v


def sanitize_node(name):
    """节点名只允许 [A-Za-z0-9_]"""
    return re.sub(r'[^A-Za-z0-9_]', '_', name)


def sanitize_signal(name):
    """信号名清洗：特殊字符 -> 下划线，合并连续下划线"""
    safe = re.sub(r'[^A-Za-z0-9_]', '_', name)
    safe = re.sub(r'_+', '_', safe).strip('_')
    return safe


def parse_val_desc(text):
    """解析 Signal Value Description 列中的枚举值定义"""
    val_defs = {}
    if not text or not text.strip():
        return val_defs
    for line in text.split('\n'):
        line = line.strip()
        # 支持 U+FF1A 全角冒号和 U+003A 半角冒号
        colon = ':'
        if '：' in line:
            colon = '：'
        if colon not in line:
            continue
        parts = line.split(colon, 1)
        val_str = parts[0].strip()
        desc_str = parts[1].strip()[:60]
        if val_str.startswith('0x') or val_str.startswith('0X'):
            try:
                val_defs[int(val_str, 16)] = desc_str
            except ValueError:
                pass
        elif val_str.isdigit():
            val_defs[int(val_str)] = desc_str
    return val_defs


def resolve_byte_order(text):
    """Byte Order 列 -> DBC byte order 标记"""
    t = text.lower()
    if 'intel' in t:
        return '@0+'  # Intel, unsigned
    else:
        return '@1+'  # Motorola, unsigned


def clean_for_ansi(text):
    """清除非 Latin-1 字符，合并连续空格"""
    cleaned = []
    for ch in text:
        if ord(ch) < 256:
            cleaned.append(ch)
        else:
            cleaned.append(' ')
    result = ''.join(cleaned)
    result = re.sub(r'  +', ' ', result)
    return result.strip()


# ════════════════════════════════════════════════════════════════
# DBC 生成
# ════════════════════════════════════════════════════════════════

def gen_dbc(frames):
    """生成 DBC 文本"""
    lines = []
    lines.append('VERSION ""')
    lines.append('')

    # NS_
    lines.append('NS_ :')
    for ns in ['NS_DESC_', 'CM_', 'BA_DEF_', 'BA_DEF_DEF_',
               'BA_DEF_DEF_REL_', 'BA_DEF_SGTYPE_', 'BA_DEF_REL_',
               'BA_DEF_DEF_', 'BU_DEF_REL_', 'BA_', 'BA_DEF_',
               'VAL_', 'CAT_DEF_', 'CAT_', 'FILTER',
               'BA_DEF_DEF_', 'EV_DATA_', 'ENVVAR_DATA_', 'SGTYPE_',
               'SGTYPE_VAL_', 'BA_DEF_SGTYPE_', 'BA_SGTYPE_', 'SIG_VALTYPE_',
               'SIGTYPE_VALTYPE_', 'BO_TX_BU_', 'BA_DEF_REL_',
               'BA_DEF_DEF_REL_', 'BA_REL_', 'BA_DEF_DEF_', 'BU_BO_REL_',
               'BA_DEF_', 'BU_SGP_REL_', 'BU_SGT_REL_', 'SG_MUL_VAL_']:
        lines.append(f'  {ns}')
    lines.append('')

    # BS_
    lines.append('BS_:')
    lines.append('')

    # BU_ - 收集所有节点
    all_nodes = set()
    for f in frames:
        if f['sender']:
            all_nodes.add(sanitize_node(f['sender']))
        for rcv in f['all_receivers']:
            all_nodes.add(sanitize_node(rcv))
    lines.append(f'BU_: {" ".join(sorted(all_nodes))}')
    lines.append('')

    # BO_ + SG_
    for f in frames:
        sigs = []
        for sig in f['signals']:
            name, start, length, factor, offset, pmin, pmax, unit, desc, bo = sig
            rcv_str = ','.join(sanitize_node(r) for r in f['all_receivers'])
            if not rcv_str:
                rcv_str = 'Vector__XXX'
            # 物理范围自动填充
            min_v = pmin
            max_v = pmax if pmax > pmin else pmin + (1 << length) - 1
            sigs.append(
                f'  SG_ {name} : {start}|{length}{bo} ({factor:g},{offset:g}) [{min_v:g}|{max_v:g}] "{unit}"  {rcv_str}'
            )
        lines.append(f'BO_ {f["id"]} {f["name"]}: {f["dlc"]} {sanitize_node(f["sender"])}')
        lines.extend(sigs)
        lines.append('')

    # CM_ SG_
    for f in frames:
        for sig in f['signals']:
            name = sig[0]
            desc = sig[8]  # English signal description (from column M)
            if desc:
                desc_clean = desc.replace('"', "'")
                lines.append(f'CM_ SG_ {f["id"]} {name} "{desc_clean}";')
    if lines[-1] != '':
        lines.append('')

    # CM_ BO_
    for f in frames:
        cmt = f.get('comment', f['name'])
        cmt_clean = cmt.replace('"', "'")
        lines.append(f'CM_ BO_ {f["id"]} "{cmt_clean}";')
    lines.append('')

    # BA_DEF_
    lines.append('BA_DEF_ BO_ "GenMsgCycleTime" INT 0 65535;')
    lines.append('BA_DEF_ BO_ "GenMsgSendType" STRING;')
    lines.append('')
    lines.append('BA_DEF_DEF_ "GenMsgCycleTime" 0;')
    lines.append('BA_DEF_DEF_ "GenMsgSendType" "Cyclic";')
    lines.append('')

    for f in frames:
        lines.append(f'BA_ "GenMsgCycleTime" BO_ {f["id"]} {f["cycle"]};')
        send_type = f.get('send_type', 'cyclic')
        if send_type:
            lines.append(f'BA_ "GenMsgSendType" BO_ {f["id"]} "{send_type}";')
    lines.append('')

    # VAL_
    for f in frames:
        fid = f['id']
        val_defs_all = f.get('val_defs_all', {})
        for sig in f['signals']:
            name = sig[0]
            bit_len = sig[2]
            if name in val_defs_all:
                vals = val_defs_all[name]
                for k in sorted(vals.keys()):
                    v_desc = vals[k].replace('"', "'")
                    lines.append(f'VAL_ {fid} {name} {k} "{v_desc}" ;')
            elif bit_len == 1:
                lines.append(f'VAL_ {fid} {name} 1 "Active" 0 "Inactive" ;')

    lines.append('')
    return '\n'.join(lines)


def save_dbc(dbc_text, output_path):
    """ANSI 编码写入"""
    dbc_clean = clean_for_ansi(dbc_text)
    with open(output_path, 'w', encoding='windows-1252') as f:
        f.write(dbc_clean)


# ════════════════════════════════════════════════════════════════
# 主转换逻辑
# ════════════════════════════════════════════════════════════════

def convert(input_path):
    """主转换流程"""
    data = parse_xlsx(input_path)

    # 1. 识别节点列 (从第1行 header 中获得)
    node_names = []
    node_cols = {}
    for col in ['AH', 'AI', 'AJ', 'AK', 'AL', 'AM']:
        header_val = get_cell(data, 1, col)
        # 提取节点名 (header 中已包含节点名)
        node_name = header_val.replace('_x000D_', '').strip()
        if node_name and re.match(r'^[A-Za-z][A-Za-z0-9_-]{0,15}$', node_name):
            node_names.append(node_name)
            node_cols[node_name] = col

    # 2. 解析行数据
    max_row = max(r for r, _ in data.keys())

    frames = []
    current = None

    for r in range(3, max_row + 1):
        msg_name = get_cell(data, r, 'A')
        sig_name = get_cell(data, r, 'G')

        if msg_name:
            # 新消息
            if current and current['signals']:
                frames.append(current)

            # 解析 ID
            id_raw = get_cell(data, r, 'C')
            can_id = 0
            if id_raw:
                id_str = id_raw.strip()
                if id_str.lower().startswith('0x'):
                    can_id = int(id_str, 16)
                elif id_str.isdigit():
                    can_id = int(id_str)

            # 周期 / DLC
            cycle_raw = get_cell(data, r, 'E')
            dlc_raw = get_cell(data, r, 'F')
            cycle = int(float(cycle_raw)) if cycle_raw and cycle_raw != '.' else 0
            dlc = int(float(dlc_raw)) if dlc_raw and dlc_raw != '.' else 8

            # 发送类型
            send_type_raw = get_cell(data, r, 'D').lower()
            if 'event' in send_type_raw:
                send_type = 'Event'
            elif 'ce' in send_type_raw or 'ca' in send_type_raw:
                send_type = 'CyclicAndEvent'
            elif 'ifactive' in send_type_raw:
                send_type = 'IfActive'
            else:
                send_type = 'Cyclic'

            # 消息类型
            msg_type = get_cell(data, r, 'B').lower()

            # 找发送/接收节点
            sender = None
            receivers = []
            for nn, nc in node_cols.items():
                v = get_cell(data, r, nc).upper()
                if v == 'TX' or v == 'S':
                    sender = nn
                elif v == 'RX' or v == 'R':
                    receivers.append(nn)

            # 查信号行中的节点分配 (NM 信号可能每行有自己的分配)
            # 对 NM 消息，从信号行收集接收者
            sig_receivers = {}

            current = {
                'name': msg_name,
                'id': can_id,
                'cycle': cycle,
                'dlc': dlc,
                'send_type': send_type,
                'msg_type': msg_type,
                'sender': sender or 'Vector__XXX',
                'all_receivers': list(dict.fromkeys(receivers)),  # 去重保序
                'signals': [],
                'val_defs_all': {},
                'comment': msg_name,
            }
            if not receivers:
                # 从消息名前缀推断
                prefix = msg_name.split('_')[0]
                for nn in node_names:
                    if sanitize_node(nn).startswith(prefix):
                        current['sender'] = nn
                        break

        if sig_name and current:
            safe_name = sanitize_signal(sig_name)

            # 读取信号字段
            start_byte_raw = get_cell(data, r, 'O')
            start_bit_raw = get_cell(data, r, 'P')
            bit_len_raw = get_cell(data, r, 'R')
            byte_order_raw = get_cell(data, r, 'N')
            factor_raw = get_cell(data, r, 'T')
            offset_raw = get_cell(data, r, 'U')
            pmin_raw = get_cell(data, r, 'V')
            pmax_raw = get_cell(data, r, 'W')
            unit_raw = get_cell(data, r, 'AC')
            val_desc_raw = get_cell_raw(data, r, 'AD')
            desc_raw = get_cell_raw(data, r, 'M')  # Signal Description

            start_byte = int(float(start_byte_raw)) if start_byte_raw and start_byte_raw != '.' else 0
            start_bit = int(float(start_bit_raw)) if start_bit_raw and start_bit_raw != '.' else 0
            bit_len = int(float(bit_len_raw)) if bit_len_raw and bit_len_raw != '.' else 1
            factor = float(factor_raw) if factor_raw and factor_raw != '.' else 1
            offset = float(offset_raw) if offset_raw and offset_raw != '.' else 0
            pmin = float(pmin_raw) if pmin_raw and pmin_raw != '.' else 0
            pmax = float(pmax_raw) if pmax_raw and pmax_raw != '.' else 255
            unit = unit_raw

            dbc_bo = resolve_byte_order(byte_order_raw)

            # 枚举值
            val_defs = parse_val_desc(val_desc_raw)
            if val_defs:
                current['val_defs_all'][safe_name] = val_defs

            # 信号行的接收节点 (NM 信号可能有单独分配)
            sig_rx = []
            for nn, nc in node_cols.items():
                v = get_cell(data, r, nc).upper()
                if v == 'RX' or v == 'R':
                    sig_rx.append(nn)

            # 如果信号行有自己的接收节点，追加到消息级接收节点
            for n in sig_rx:
                if n not in current['all_receivers']:
                    current['all_receivers'].append(n)

            # 使用英文描述作为 CM_ SG_ 注释
            eng_desc = desc_raw.replace('_x000D_', '').replace('\n', ' ').replace('\r', ' ').strip()
            # 清理掉中文等非 ASCII 字符
            eng_desc_clean = clean_for_ansi(eng_desc)

            current['signals'].append(
                (safe_name, start_bit, bit_len, factor, offset,
                 pmin, pmax, unit, eng_desc_clean, dbc_bo)
            )

    if current and current['signals']:
        frames.append(current)

    return frames


def print_summary(frames, output_path):
    """打印转换摘要"""
    total_sigs = sum(len(f['signals']) for f in frames)
    print(f'DBC 文件: {output_path}')
    print(f'帧数: {len(frames)}')
    print(f'信号数: {total_sigs}')
    print()
    print(f'{"ID":>8s}  {"Frame Name":35s} {"DLC":4s} {"Cycle":6s}  {"Direction":20s}  {"Signals":8s}')
    print('-' * 90)
    for f in frames:
        rcv_str = ','.join(f['all_receivers']) if f['all_receivers'] else 'None'
        direction = f'{f["sender"]:8s} -> {rcv_str:12s}'
        print(f'  0x{f["id"]:03X}  {f["name"]:35s} {f["dlc"]:4d} {f["cycle"]:5d}ms  {direction:20s}  {len(f["signals"]):3d}')
    print()


# ════════════════════════════════════════════════════════════════
# 入口
# ════════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) < 2:
        print('用法: python cx835ex2dbc.py <input.xlsx> [-o output.dbc]')
        sys.exit(1)

    input_path = sys.argv[1]
    if '-o' in sys.argv:
        idx = sys.argv.index('-o')
        output_path = sys.argv[idx + 1]
    else:
        base, _ = os.path.splitext(input_path)
        output_path = base + '.dbc'

    print(f'输入: {input_path}')
    print(f'输出: {output_path}')
    print()

    frames = convert(input_path)

    if not frames:
        print('错误: 没有解析到任何帧')
        sys.exit(1)

    # 生成所有节点的集合，用于 BU_ 声明
    all_nodes = set()
    for f in frames:
        if f['sender']:
            all_nodes.add(sanitize_node(f['sender']))
        for rcv in f['all_receivers']:
            all_nodes.add(sanitize_node(rcv))

    dbc_text = gen_dbc(frames)
    save_dbc(dbc_text, output_path)
    print_summary(frames, output_path)
    print('生成完成!')


if __name__ == '__main__':
    main()
