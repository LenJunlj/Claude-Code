#!/usr/bin/env python3
"""
excel-to-ldf — 从 LIN 通信矩阵 Excel 生成 Vector LDF Explorer Pro 兼容的 LDF 文件

依赖: pip install openpyxl
用法: python excel_to_ldf.py <输入Excel.xlsx> [输出LDF.ldf]
"""
import sys, re, openpyxl

def excel_to_ldf(excel_path, output_path=None):
    """主函数：读取 Excel，写入 LDF"""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    if output_path is None:
        output_path = excel_path.rsplit('.', 1)[0] + '.ldf'
    
    # ==================== 解析 Info sheet ====================
    ws_info = wb['Info']
    baudrate = str(ws_info.cell(row=3, column=2).value or '19.2').strip()
    time_base = str(ws_info.cell(row=3, column=3).value or '5')
    jitter = str(ws_info.cell(row=3, column=4).value or '0.1')
    
    master_name = None
    slaves = []
    for r in range(7, ws_info.max_row + 1):
        name = ws_info.cell(row=r, column=1).value
        nad = ws_info.cell(row=r, column=2).value
        if name:
            name = str(name).strip()
            nad_str = str(nad or '').strip()
            if nad_str == '-' or nad_str == '':
                master_name = name
            else:
                slaves.append((name, nad_str))
    
    # ==================== 解析 Matrix sheet ====================
    ws = wb['Matrix']
    col = {
        'MsgID': 1, 'MsgName': 4, 'MsgLength': 6,
        'SignalName': 12, 'StartBit': 16, 'BitLength': 17, 'InitValue': 26,
        'Resolution': 20, 'Offset': 21, 'MinPhys': 22, 'MaxPhys': 23,
        'Unit': 28, 'ValueDescription': 29,
        'MasterCol': 34, 'SlaveCol': 35
    }
    
    messages = []
    cur = None
    for r in range(2, ws.max_row + 1):
        mid = ws.cell(row=r, column=col['MsgID']).value
        if mid is not None:
            mid_str = str(mid).strip()
            if not (mid_str.startswith('0x') or mid_str.startswith('0X')):
                mid_str = hex(int(mid_str))
            cur = {
                'id': mid_str,
                'name': ws.cell(row=r, column=col['MsgName']).value,
                'length': int(ws.cell(row=r, column=col['MsgLength']).value or 8),
                'master_dir': str(ws.cell(row=r, column=col['MasterCol']).value or '').strip().lower(),
                'slave_dir': str(ws.cell(row=r, column=col['SlaveCol']).value or '').strip().lower(),
                'signals': []
            }
            messages.append(cur)
        
        sname = ws.cell(row=r, column=col['SignalName']).value
        if sname and cur:
            res_val = ws.cell(row=r, column=col['Resolution']).value
            off_val = ws.cell(row=r, column=col['Offset']).value
            sig = {
                'name': str(sname).strip(),
                'start_bit': ws.cell(row=r, column=col['StartBit']).value,
                'bit_length': ws.cell(row=r, column=col['BitLength']).value,
                'init_value': ws.cell(row=r, column=col['InitValue']).value,
                'resolution': float(res_val) if res_val is not None else 1.0,
                'offset': int(float(off_val)) if off_val is not None else 0,
                'min_phys': ws.cell(row=r, column=col['MinPhys']).value,
                'max_phys': ws.cell(row=r, column=col['MaxPhys']).value,
                'value_description': ws.cell(row=r, column=col['ValueDescription']).value,
            }
            cur['signals'].append(sig)
    
    # ==================== 解析 Schedule ====================
    ws_sched = wb['LIN Schedule']
    schedule = []
    for r in range(5, ws_sched.max_row + 1):
        mid = ws_sched.cell(row=r, column=2).value
        delay = ws_sched.cell(row=r, column=3).value
        if mid is not None:
            schedule.append((str(mid).strip(), int(delay)))
    
    # ==================== 辅助函数 ====================
    def init_val_str(val):
        """init_value 必须用十进制"""
        if val is None:
            return '0'
        s = str(val).strip()
        if s.startswith('0x') or s.startswith('0X'):
            return str(int(s, 16))
        try:
            return str(int(s))
        except:
            return s
    
    def get_pub_sub(msg):
        """根据 Excel 中 master/slave 列的 s/r 确定发布者和订阅者"""
        if msg['master_dir'] == 's' and msg['slave_dir'] == 'r':
            return master_name, slaves[0][0] if slaves else 'SRF'
        elif msg['master_dir'] == 'r' and msg['slave_dir'] == 's':
            return slaves[0][0] if slaves else 'SRF', master_name
        return master_name, slaves[0][0] if slaves else 'SRF'
    
    def parse_value_desc(text):
        """解析 '0x0:desc\\n0x1:desc2...' 为 [(int_val, desc), ...]"""
        if not text:
            return []
        result = []
        for line in text.strip().split('\n'):
            line = line.strip()
            if not line:
                continue
            if '~' in line:
                rp, _, desc = line.partition(':')
                ss, _, es = rp.partition('~')
                try:
                    si = int(ss.strip(), 16)
                    ei = int(es.strip(), 16)
                except:
                    continue
                for v in range(si, ei + 1):
                    result.append((v, desc.strip()))
            elif ':' in line:
                vs, desc = line.split(':', 1)
                try:
                    v = int(vs.strip(), 16)
                    result.append((v, desc.strip()))
                except:
                    continue
        return result
    
    # ==================== 构建 LDF 内容 ====================
    lines = []
    
    # 开头两个空行
    lines.append("")
    lines.append("")
    lines.append("LIN_description_file;")
    lines.append('LIN_protocol_version = "2.0";')   # 固定 2.0
    lines.append('LIN_language_version = "2.0";')
    lines.append(f'LIN_speed = {baudrate} kbps;')
    lines.append("")
    
    # Nodes
    lines.append("Nodes {")
    lines.append(f"  Master: {master_name}, {time_base} ms, {jitter} ms ;")
    if slaves:
        lines.append(f"  Slaves: {', '.join(s[0] for s in slaves)} ;")
    lines.append("}")
    lines.append("")
    
    # Signals — 格式: name: length, init_value(十进制), publisher, subscriber ;
    lines.append("Signals {")
    for m in messages:
        pub, sub = get_pub_sub(m)
        for sig in m['signals']:
            iv = init_val_str(sig['init_value'])
            lines.append(f"  {sig['name']}: {sig['bit_length']}, {iv}, {pub}, {sub} ;")
    lines.append("}")
    lines.append("")
    
    # Diagnostic_signals
    lines.append("Diagnostic_signals {")
    for i in range(8):
        lines.append(f"  MasterReqB{i}: 8, 0 ;")
    for i in range(8):
        lines.append(f"  SlaveRespB{i}: 8, 0 ;")
    lines.append("}")
    lines.append("")
    
    # Frames — ID 用十进制
    lines.append("Frames {")
    for m in messages:
        pub, _ = get_pub_sub(m)
        fid_dec = int(m['id'], 16)
        lines.append(f"  {m['name']}: {fid_dec}, {pub}, {m['length']} {{")
        for sig in m['signals']:
            lines.append(f"    {sig['name']}, {sig['start_bit']} ;")
        lines.append("  }")
    lines.append("}")
    lines.append("")
    
    # Diagnostic_frames
    lines.append("Diagnostic_frames {")
    lines.append("  MasterReq: 0x3c {")
    for i in range(8):
        lines.append(f"    MasterReqB{i}, {i*8} ;")
    lines.append("  }")
    lines.append("  SlaveResp: 0x3d {")
    for i in range(8):
        lines.append(f"    SlaveRespB{i}, {i*8} ;")
    lines.append("  }")
    lines.append("}")
    lines.append("")
    
    # Node_attributes
    lines.append("Node_attributes {")
    for sname, nad in slaves:
        lines.append(f"  {sname}{{")    # 大括号在同一行
        lines.append(f'    LIN_protocol = "2.0" ;')
        lines.append(f"    configured_NAD = {nad} ;")
        lines.append(f"    product_id = 0x0, 0x0, 255 ;")
        lines.append(f"    P2_min = 0 ms ;")
        lines.append(f"    ST_min = 0 ms ;")
        lines.append(f"    configurable_frames {{")
        for m in messages:
            lines.append(f"      {m['name']} = 0x0 ;")
        lines.append(f"    }}")
        lines.append(f"  }}")
    lines.append("}")
    lines.append("")
    
    # Schedule_tables
    lines.append("Schedule_tables {")
    lines.append(f" {master_name}_Schedule {{")    # 缩进一个空格
    for mid_str, delay in schedule:
        fname = None
        for m in messages:
            if m['id'].lower() == mid_str.lower():
                fname = m['name']
                break
        if fname:
            lines.append(f"    {fname} delay {delay} ms ;")
    lines.append("  }")
    lines.append("}")
    lines.append("")
    
    # Signal_encoding_types
    lines.append("Signal_encoding_types {")
    for m in messages:
        for sig in m['signals']:
            vd = sig['value_description']
            has_phys = (sig['resolution'] != 1.0 or sig['offset'] != 0)
            if vd or has_phys:
                lines.append(f"  {sig['name']}_Encoding {{")
                
                # logical_value 条目（来自值描述）
                if vd:
                    for val, desc in parse_value_desc(vd):
                        if desc:
                            lines.append(f'    logical_value, {val}, "{desc}" ;')
                        else:
                            lines.append(f'    logical_value, {val} ;')
                
                # physical_value 条目（分辨率/偏移非默认）
                if has_phys:
                    min_p = float(sig['min_phys']) if sig['min_phys'] is not None else 0.0
                    max_p = float(sig['max_phys']) if sig['max_phys'] is not None else 0.0
                    # min/max 必须是 raw 值（整数）！
                    min_raw = int(min_p / sig['resolution'] - sig['offset'])
                    max_raw = int(max_p / sig['resolution'] - sig['offset'])
                    lines.append(f"    physical_value, {min_raw}, {max_raw}, {sig['resolution']}, {sig['offset']} ;")
                    # physical_value 不能单独存在，必须搭配 logical_value
                    if not vd:
                        lines.append(f'    logical_value, {min_raw}, "{min_p}" ;')
                        lines.append(f'    logical_value, {max_raw}, "{max_p}" ;')
                
                lines.append(f"  }}")
    lines.append("}")
    lines.append("")
    
    # Signal_representation — 格式: encoding_name : signal_name ;
    lines.append("Signal_representation {")
    for m in messages:
        for sig in m['signals']:
            if sig['value_description'] or sig['resolution'] != 1.0 or sig['offset'] != 0:
                lines.append(f"  {sig['name']}_Encoding: {sig['name']} ;")
    lines.append("}")
    lines.append("")
    
    # ==================== 写入文件 ====================
    content = '\n'.join(lines)
    
    # UTF-8 without BOM + CRLF
    with open(output_path, 'w', encoding='utf-8', newline='\r\n') as f:
        f.write(content)
    
    # ==================== 验证 ====================
    opens = content.count('{')
    closes = content.count('}')
    balanced = opens == closes
    
    sig_count = sum(len(m['signals']) for m in messages)
    enc_count = sum(1 for m in messages for s in m['signals']
                    if s['value_description'] or s['resolution'] != 1.0 or s['offset'] != 0)
    
    print(f"[OK] LDF: {output_path}")
    print(f"   signals={sig_count}, frames={len(messages)}, encodings={enc_count}")
    print(f"   brackets: {{={opens}, }}={closes}, balanced={balanced}")
    
    # 列出没有 encoding 的信号（供参考）
    no_enc = [s['name'] for m in messages for s in m['signals']
              if not s['value_description'] and s['resolution'] == 1.0 and s['offset'] == 0]
    if no_enc:
        print(f"   无编码: {', '.join(no_enc)}")
    
    return output_path

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python excel_to_ldf.py <输入Excel.xlsx> [输出LDF.ldf]")
        sys.exit(1)
    excel_to_ldf(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
