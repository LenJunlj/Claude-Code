---
name: excel-to-dbc
description: >
  Use when the user asks to generate a CANoe DBC file from a communication
  matrix Excel (.xlsx) file, or when working with CAN bus signal matrices.
  Converts Excel communication matrices (with Msg ID, Signal Name, Start Byte,
  Start Bit, Bit Length, Byte Order columns) into ANSI-encoded DBC files.
  Handles Motorola/Intel byte order, value enumerations (VAL_), signal comments
  (CM_), message cycle times (BA_), and signal layout mapping.
  Triggers on: generate DBC, convert Excel to DBC, communication matrix to DBC,
  生成DBC, 导出DBC, DBC文件.
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# Excel 通信矩阵 → CANoe DBC 生成规则

从通信矩阵 Excel 生成 CANoe DBC 文件时的关键规则和踩坑记录。所有规则基于 `scripts/signal2dbc.py` 的实践经验总结。

---

## 工作流程

1. **解析 Excel**：读取 Matrix sheet，Msg Name 列有值的行为新报文起始行，无 Msg Name 但有 Signal Name 的行为信号行
2. **提取元数据**：从报文行获取 Msg ID / Msg Type / Msg Send Type / Msg Cycle Time / Msg Length / 发送方与接收方
3. **提取信号**：从信号行获取 Signal Name / Byte Order / Start Byte / Start Bit / Bit Length / Data Type / Resolution / Offset / Min / Max / Unit / Value Description
4. **生成 DBC**：按下面规则生成 BO_ / SG_ / CM_ / BA_ / VAL_ 块
5. **保存文件**：ANSI (Windows-1252) 编码写入

## 输入格式

Excel 的 `Matrix` sheet 包含以下列：

| 列 | 说明 |
|---|---|
| Msg Name | 报文名称，如 `DSMC_0x325` |
| Msg Type | Normal / NM / Diag |
| Msg ID | 十六进制 CAN ID，如 `0x325` |
| Msg Send Type | Cycle / Event / CE / CA / IfActive |
| Msg Cycle Time(ms) | 周期时间 |
| Msg Length(Byte) | DLC |
| Signal Name | 信号名称 |
| Byte Order | Motorola LSB / Intel |
| Start Byte / Start Bit / Bit Length / End Bit | 信号布局 |
| Data Type | Unsigned / Signed |
| Resolution / Offset | 物理值转换因子 |
| Signal Min/Max Value (phys / hex) | 物理范围与原始值范围 |
| Initial Value(Hex) / Invalid Value(Hex) | 初始值与无效值 |
| Unit | 单位 |
| Signal Value Description | 枚举值描述，格式 `0x0: Description` |
| ECU 列 (DSMC, BDCU, ...)  | `S` = 发送方, `R` = 接收方 |

---

## 1. 文件编码

- **保存为 ANSI (Windows-1252) 编码**
- CANoe 对 DBC 文件期望 ANSI 编码，UTF-8 会导致解析失败
- 所有非 Latin-1 字符（中文、特殊符号）必须清除或替换为空格
- 保存后合并连续空格：`re.sub(r'  +', ' ', dbc_clean)`

## 2. CM_ 注释

- **所有注释必须为纯英文**，不可包含中文
- 包括 `CM_ SG_` 和 `CM_ BO_` 注释
- **禁止写入 `CM_ SG_ <id> <name> "";`** — 该语法在某些 DBC 解析器中不被识别，会引发解析异常
- 信号注释取 English Signal Description 列，若无则取 Chinese Signal Description（但要注意 ANSI 编码下中文会被清掉）
- 注释中的双引号替换为单引号

## 3. 节点命名

- DBC 节点名只允许 `[A-Za-z0-9_]`
- **连字符 `-` 必须替换为下划线 `_`**，例如 `PAD-1` → `PAD_1`
- 节点名同时出现在 `BU_:` 声明和信号接收节点字段，两处都要清洗
- 发送方从 ECU 列中标记为 `S` 的节点获取；若无标记，从报文名前缀推断

## 4. "Start Bit" 列含义（最容易踩坑）

- Excel 的 "Start Bit" 列使用 **LSB0 编号**（LSB = 0）
- Bit=0, Len=2 表示信号占用 **bit0 和 bit1**
- **不需要做任何转换**，直接作为 DBC 的 start_bit

### 历史踩坑记录

| 尝试 | 公式 | 结果 |
|------|------|------|
| ❌ MSB0转换 | `start = byte×8 + 7 − (bit−byte×8)` | 起始位算反（7→0错位） |
| ❌ 转MSB位置 | `start = bit + len − 1` | Motorola下偏移了1位 |
| ✅ 直接用 | `start = bit` | 与Excel完全一致 |

## 5. 字节序

- **按矩阵的 "Byte Order" 列填写**
- 矩阵标 Motorola → DBC 用 `@1+`  (Motorola / big-endian)
- 矩阵标 Intel → DBC 用 `@0+`  (Intel / little-endian)
- 不要因为位编号是 LSB0 就擅自改用 Intel

## 6. 信号命名

- 信号名只允许 `[A-Za-z0-9_]`
- Excel 中的特殊字符（空格、括号、斜杠等）替换为下划线
- 连续多个下划线合并为一个：`re.sub(r'_+', '_', safe_name)`
- 首尾下划线去除：`.strip('_')`

## 7. VAL_ 枚举值

- 从 Excel "Signal Value Description" 列解析
- 支持 `0x0: xxx` 和 `0: xxx` 两种格式
- 枚举描述只保留前 60 个字符，双引号替换为单引号
- 对于 1-bit 信号，自动添加 `Active`/`Inactive` 枚举
- VAL_ 描述中的中文同样会被 ANSI 编码清洗

## 8. 物理值范围

- 如 Signal Max Value(phys) ≤ Signal Min Value(phys)，自动补齐为 `pmin + (1 << bit_len) - 1`
- Resolution 和 Offset 按整数格式显示（去掉多余小数位）

## 9. DBC 结构模板

```
VERSION ""

NS_ :
    NS_DESC_
    CM_
    BA_DEF_
    BA_
    VAL_
    ...

BS_:

BU_: <发送节点> <接收节点1> <接收节点2> ...

BO_ <CAN_ID> <MsgName>: <DLC> <Sender>
 SG_ <SignalName> : <StartBit>|<BitLen>@<ByteOrder> (<Factor>,<Offset>) [<Min>|<Max>] "<Unit>" <Receiver>

CM_ SG_ <CAN_ID> <SignalName> "<EnglishDescription>";
CM_ BO_ <CAN_ID> "<MsgName>";

BA_DEF_ BO_ "GenMsgCycleTime" INT 0 65535;
BA_DEF_ BO_ "GenMsgSendType" STRING ;
BA_DEF_DEF_ "GenMsgCycleTime" 0;
BA_DEF_DEF_ "GenMsgSendType" "Cyclic";
BA_ "GenMsgCycleTime" BO_ <CAN_ID> <CycleMs>;
BA_ "GenMsgSendType" BO_ <CAN_ID> <SendType>;

VAL_ <CAN_ID> <SignalName> <Value> "<Description>" ;
```

## 10. DBC 信号物理布局验证（自检清单）

生成 DBC 后检查以下几点可以提前发现问题：

- [ ] 起始位是否和 Excel "Start Bit" 列数值一致
- [ ] 字节序标记是否和 Excel "Byte Order" 列一致
- [ ] 连续排列的 2-bit 信号是否紧贴无间隙（Bit 0→2→4→6→8...）
- [ ] 所有节点名不含连字符
- [ ] 零非 ASCII 字节
- [ ] 无 `CM_ SG_ <id> <name> "";` 空注释行
- [ ] 报文数与信号数与 Excel 一致
- [ ] VAL_ 枚举值正确对应
