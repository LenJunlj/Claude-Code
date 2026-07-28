---
name: can-test-tool
description: |
  CAN Test Tool for FSCM/RSCM — a tkinter-based GUI application that parses DBC files,
  provides transmit signal controls and receive signal displays over PCAN, Vector/CANoe,
  or virtual CAN interfaces. Supports CAN FD with configurable bitrates.

allowed-tools:
  - Read
  - Write
  - Edit
  - Glob
  - Grep
  - Bash
  - WebSearch

trigger-phrases:
  - CAN test tool
  - FSCM RSCM
  - DBC parser
  - CAN communication
  - generate DBC from Excel
  - 上位机
  - CAN工具
  - 通信矩阵
  - DBC文件
  - CAN FD
  - CANoe
  - PCAN
  - Vector
---

# CAN Test Tool — Claude Code Skill

## 1. 设计原则

### 1.1 数据源优先级
- **通信矩阵（Excel）为权威主数据源**，DBC文件由矩阵自动生成（副）
- 所有信号定义以Excel `FSCM_LDCANFD` sheet为准
- DBC文件仅为运行时解析用，不手工编辑
- 更新流程：更新Excel → 运行`generate_dbc_from_excel.py` → 重启上位机

### 1.2 字节序规范
- **所有信号使用 Motorola MSB（Big Endian）**，byte_order=1
- 常见错误：DBC文件用Intel（Little Endian）但硬件发Motorola，导致bit位错乱
- Excel中Byte Order列值为"Motorola MSB"

---

## 2. 项目结构

```
panel/CAN_TestTool/
├── can_test_tool.py              # 主GUI应用（tkinter，含main入口）
├── dbc_parser.py                 # DBC文件解析器
├── can_comms.py                  # CAN通信接口层（PCAN/Vector/Virtual）
├── generate_dbc_from_excel.py    # Excel → DBC 生成器
├── launch.bat                    # Windows启动脚本
├── SKILL.md                      # 本文档（Claude Code Skill）
└── CAN_TestTool_Skills.md        # 旧版知识文档（待迁移）

panel/
├── CX835_VP2_LDCAN_FSCM_20250211_Fix.dbc    # DBC文件
└── CX835EX_VP_LDCANFD_FSCM_20251216_Fix.xlsx # Excel通信矩阵
```

---

## 3. DBC解析器（dbc_parser.py）

### 3.1 核心数据结构

```python
@dataclass
class Signal:
    name: str
    message_id: int
    start_bit: int          # Motorola: MSB的flat bit position
    size: int               # 信号bit长度
    byte_order: int         # 0=Intel, 1=Motorola
    signed: bool
    factor: float           # 分辨率/精度
    offset: float           # 偏移量
    minimum: float          # 物理最小值
    maximum: float          # 物理最大值
    unit: str               # 单位
    receivers: List[str]    # 接收节点列表
    value_descriptions: Dict[int, str]  # 枚举值描述 {raw: text}
    gen_sig_start_value: Optional[float]  # 初始值（已转为物理值）
    gen_sig_invalid_value: Optional[int]  # 无效值（raw）
    gen_sig_send_type: int  # 信号发送类型

@dataclass
class Message:
    id: int
    name: str
    dlc: int                # 数据长度码
    transmitter: str        # 发送节点
    signals: Dict[str, Signal]
    cycle_time: int         # 周期时间(ms)
    send_type: int          # 0=Cycle, 1=NoMsgSendType, 2=IfActive, 3=Event, 4=CA, 5=CE
    is_nm: bool             # Network Management消息
    is_diag_request: bool
    is_diag_response: bool
```

### 3.2 信号编码/解码算法

**转换公式**：`PhysicalValue = RawValue × factor + offset`

#### 解码（CAN bytes → 物理值）
```python
def decode(data: bytes) -> float:
    raw = self._extract_raw(data)  # 根据byte_order提取
    return raw * self.factor + self.offset
```

#### 编码（物理值 → CAN bytes）
```python
def encode(physical_value: float) -> int:
    raw = int((physical_value - self.offset) / self.factor + 0.5)
    return clamp(raw, 0, (1 << size) - 1)
```

### 3.3 Motorola MSB 位提取（关键算法）

Motorola MSB中，start_bit是MSB的flat bit position：
- flat bit = byte × 8 + bit_within_byte
- bit_within_byte 从0（LSB）到7（MSB）

```python
def _extract_motorola(data: bytes) -> int:
    raw = 0
    for i in range(self.size):           # i=0 → LSB of raw
        msb_bit = self.start_bit         # MSB flat position
        bit_pos = msb_bit - i            # 从MSB向下走
        byte_idx = bit_pos // 8
        bit_in_byte = 7 - (bit_pos % 8)  # 反转：flat bit 7→byte bit 0
        if data[byte_idx] & (1 << bit_in_byte):
            raw |= (1 << i)              # 放入raw的对应位
    return raw
```

#### Motorola pack（编码到bytes）
```python
def pack(self, physical_value: float, data: bytearray):
    raw = self.encode(physical_value)
    for i in range(self.size):
        msb_bit = self.start_bit
        bit_pos = msb_bit - i
        byte_idx = bit_pos // 8
        bit_in_byte = 7 - (bit_pos % 8)
        if raw & (1 << i):
            data[byte_idx] |= (1 << bit_in_byte)
        else:
            data[byte_idx] &= ~(1 << bit_in_byte)
```

### 3.4 Intel 位提取

Intel中，start_bit是LSB的flat bit position，向上递增：
```python
def _extract_intel(data: bytes) -> int:
    raw = 0
    for i in range(self.size):
        bit_pos = self.start_bit + i    # 从LSB向上走
        byte_idx = bit_pos // 8
        bit_in_byte = bit_pos % 8
        if data[byte_idx] & (1 << bit_in_byte):
            raw |= (1 << i)
    return raw
```

### 3.5 DBC解析关键

| 正则模式 | 用途 |
|----------|------|
| `BO_ (id) (name) : (dlc) (transmitter)` | 消息定义 |
| `SG_ (name) : (start_bit)\|(size)@(byte_order)(signed)` | 信号定义 |
| `CM_ SG_ (msg_id) (sig_name) "comment";` | 信号注释 |
| `VAL_ (msg_id) (sig_name) val "text" ... ;` | 枚举值描述 |
| `BA_ "GenMsgCycleTime" BO_ (msg_id) (ms);` | 周期时间 |
| `BA_ "GenMsgSendType" BO_ (msg_id) (type);` | 发送类型 |
| `BA_ "NmMessage" BO_ (msg_id) 1;` | NM消息标记 |
| `BA_ "GenSigStartValue" SG_ (msg_id) (sig_name) (raw);` | 信号初值 |
| `BA_ "GenSigInvalidValue" SG_ (msg_id) (sig_name) (raw);` | 信号无效值 |

### 3.6 消息分类（classify_messages）

```python
def classify_messages(db):
    to_fscm = {}    # 发送给FSCM的（左面板控件）
    from_fscm = {}  # 从FSCM接收的（右面板显示）
    for msg_id, msg in db.messages.items():
        if msg.is_nm:
            continue
        if msg.transmitter == 'FSCM':
            from_fscm[msg_id] = msg     # FSCM发出的
        elif 'FSCM' in any_signal_receivers:
            to_fscm[msg_id] = msg       # FSCM接收的
    return to_fscm, from_fscm
```

---

## 4. CAN通信层（can_comms.py）

### 4.1 接口层次

```
CanInterface (抽象基类)
├── PcanInterface     # PCAN-USB硬件（python-can pcan backend）
├── VirtualCanInterface  # 虚拟环回（无硬件测试）
└── CanoeInterface   # Vector硬件 VN1610等（python-can vector backend）

CanManager           # 管理器（封装接口生命周期）
```

### 4.2 CanInterface 基类

```python
class CanInterface:
    _rx_callbacks: Dict[int, list]  # arb_id → [callbacks]
    _rx_all_callbacks: list         # 所有消息的回调
    _running: bool                  # RX线程标志
    _rx_thread: Optional[Thread]

    def register_rx_callback(arb_id, callback)     # 注册ID过滤回调
    def register_rx_all_callback(callback)          # 注册全量回调
    def _notify_rx(msg):                            # 分发消息到回调
```

### 4.3 PcanInterface 关键实现

```python
class PcanInterface(CanInterface):
    def open(self, channel='PCAN_USBBUS1', bitrate=500000,
             fd_mode=False, data_bitrate=2000000):
        import can
        if channel.isdigit():   # 支持数字通道号如"82"
            chan = channel
        else:
            chan = self._channel_map.get(channel, channel)
        kwargs = dict(bustype='pcan', channel=chan, bitrate=bitrate)
        if fd_mode:
            kwargs['fd'] = True
            kwargs['data_bitrate'] = data_bitrate
        self._bus = can.interface.Bus(**kwargs)
        # 启动RX线程...

    def _rx_loop(self):
        while self._running:
            msg = self._bus.recv(timeout=0.05)
            if msg is not None:
                cmsg = CanMessage(
                    arb_id=msg.arbitration_id,
                    data=bytes(msg.data) if msg.data else b'',
                    dlc=msg.dlc,
                )
                self._notify_rx(cmsg)
```

### 4.4 通道号支持

| 输入 | 映射 |
|------|------|
| `"PCAN_USBBUS1"` | PCAN_USBBUS1（命名通道） |
| `"82"` | 82（数字通道，直接传递） |
| `"virtual"` | VirtualCanInterface |
| `"CANoe"` | CanoeInterface（Vector硬件） |

### 4.5 CanManager API

```python
mgr = CanManager()
mgr.list_available_devices()           # 静态方法：扫描PCAN+Vector设备
mgr.connect_pcan(channel, bitrate, fd_mode, data_bitrate)  # 连接PCAN
mgr.connect_virtual()                  # 虚拟模式
mgr.connect_canoe(channel, bitrate, fd_mode, data_bitrate) # Vector硬件+CAN FD
mgr.disconnect()
mgr.send_message(arb_id, data)
mgr.register_rx_callback(arb_id, cb)   # ID过滤回调
mgr.register_rx_all_callback(cb)       # 全量回调
mgr.start_cyclic_send(arb_id, data_func, interval_ms)
mgr.stop_cyclic_send(arb_id)
mgr.stop_all_cyclic()
```

### 4.6 设备自动扫描

```python
@staticmethod
def list_available_devices() -> list:
    import can
    devices = []
    # Scan PCAN
    for cfg in can.detect_available_configs(bustype='pcan'):
        ch = cfg.get('channel', '')
        if ch.startswith('pcan:'):
            ch = ch[5:]
        devices.append(ch)
    # Scan Vector
    for cfg in can.detect_available_configs(bustype='vector'):
        ch = cfg.get('channel', '')
        if ch and ch not in devices:
            devices.append(ch)
    # 确保标准PCAN通道+virtual
    devices.append('virtual')
    return devices
```

### 4.7 RX回调修复（关键）

RX回调必须在`connect_*()`之前注册，否则会丢失消息。修复方案：

```python
class CanManager:
    def __init__(self):
        self._pending_rx_all_callbacks: list = []
        self._pending_rx_callbacks: Dict[int, list] = {}

    def register_rx_all_callback(self, callback):
        # 无论是否已连接，先存储
        self._pending_rx_all_callbacks.append(callback)
        if self._interface:
            self._interface.register_rx_all_callback(callback)

    def _flush_pending_rx_callbacks(self):
        """connect成功后重注册所有pending回调"""
        for cb in self._pending_rx_all_callbacks:
            self._interface.register_rx_all_callback(cb)
        for arb_id, cbs in self._pending_rx_callbacks.items():
            for cb in cbs:
                self._interface.register_rx_callback(arb_id, cb)
```

### 4.8 线程模型

- 每个接口在`open()`时启动一个**daemon RX线程**
- RX线程循环调用`bus.recv(timeout=0.05)` 轮询
- 收到消息 → `_notify_rx()` → 分发给所有注册回调
- **GUI更新必须使用`self.after(0, callback)`** 委托到主线程（tkinter线程安全约束）
- 数据访问（如`_rx_count += 1`）在RX线程直接操作，利用GIL保证原子性

---

## 5. GUI层（can_test_tool.py）

### 5.1 整体布局

```
CanTestTool (tk.Tk)
├── Menu Bar               # File(Load DBC/Exit) / CAN(Connect/Disconnect/Cyclic) / Help
├── Toolbar (2行)
│   ├── Row 0: Ch: [channel_combo] ⟳  Mode:[CAN|CAN FD]  Baud:[125-1000]  Data:[500-8000]
│   └── Row 1: [Connect] [Disconnect] | [Cyclic Send] [Send All Once]
├── Main Area (PanedWindow 左右分栏)
│   ├── Left: SendPanel (发送给FSCM的16条消息)
│   │   └── ScrollableFrame → SendMessagePanel × N (按Category分组)
│   └── Right: RecvPanel (从FSCM接收+动态创建)
│       └── ScrollableFrame → RecvMessagePanel × N
└── Status Bar
    ├── Status: Connected/Disconnected
    ├── TX: count
    └── RX: count
```

### 5.2 CAN FD工具栏配置

```
Ch: [PCAN_USBBUS1 ▼] [⟳]  Mode: [CAN ▼]  Baud: [500 ▼]  Data: [2000 ▼]  (Data仅在FD模式可用)

[Connect] [Disconnect]  |  [☐ Cyclic Send] [Send All Once]
```

- `Mode: CAN/CAN FD` 切换时自动启用/禁用Data rate下拉框
- `Baud` 为仲裁段波特率（kbaud），`Data` 为数据段波特率（kbaud，仅CAN FD）

### 5.3 可滚动框架

```python
class ScrollableFrame(ttk.Frame):
    """Canvas + Scrollbar实现的双向可滚动区域"""
    def __init__(self, parent):
        self.canvas = tk.Canvas(...)
        self.v_scrollbar = ttk.Scrollbar(orient='vertical')
        self.h_scrollbar = ttk.Scrollbar(orient='horizontal')
        self.inner_frame = ttk.Frame(self.canvas)  # 实际内容放在这
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor='nw')
        # 鼠标滚轮：bind_all('<MouseWheel>')
```

### 5.4 信号控件类型

| 控件类 | 适用场景 | 组件 |
|--------|----------|------|
| `AnalogSignalWidget` | 模拟量（size>1且非枚举） | Scale滑块 + 数值Label |
| `EnumSignalWidget` | 枚举值（有value_descriptions） | Combobox下拉 |
| `BinarySignalWidget` | 开关量（size=1） | Combobox [0, 1] |

所有控件统一接口：
```python
widget.get_value() -> float
widget.set_value(val, notify=True)
```

### 5.5 发送面板（SendMessagePanel）

- **LabelFrame标题**：`0x{ID} {Name} [{SendType}]`
- Enable复选框：控制该消息是否参与发送
- 信号按`start_bit`排序后2列grid布局
- 带Seat分组（LFSeat/RFSeat/LRSeat/RRSeat/MirrorLeft/MirrorRight）
- 底部Button Row：`Send Once` + `Cyclic OFF/ON`
- 初始值从`GenSigStartValue`属性读取

信号分组：
```python
prefixes = ['LFSeat', 'RFSeat', 'LRSeat', 'RRSeat', 'MirrorLeft', 'MirrorRight']
# 按前缀分组显示，无前缀的归为"Other"
```

### 5.6 接收面板（RecvMessagePanel）

- **LabelFrame标题**：`0x{ID} {Name} (Tx: {transmitter})`
- 信号2列grid布局
- 每次更新闪烁效果（foreground变蓝200ms后恢复）
- 值不变时不更新UI（通过`_last_values`对比）

### 5.7 RecvPanel动态创建

```python
class RecvPanel(ttk.Frame):
    """接收面板，支持动态创建未预建的消息面板"""
    def __init__(self, parent, from_fscm, can_mgr, db_messages=None):
        # 预建所有from_fscm消息的面板
        for msg_id in sorted(from_fscm.keys()):
            self._add_panel(from_fscm[msg_id])

    def update_message(self, msg_id, data):
        """更新或动态创建消息面板"""
        if msg_id in self._msg_displays:
            self._msg_displays[msg_id].update_from_data(data)
        elif msg_id in self._db_messages:
            # 收到未预建的消息 → 动态创建面板
            self._add_panel(self._db_messages[msg_id])
            self._msg_displays[msg_id].update_from_data(data)
```

### 5.8 RX消息处理

```python
def _on_rx_message(self, msg: CanMessage):
    """从RX线程回调，由CanManager调用"""
    self._rx_count += 1
    self.after(0, lambda: self.rx_count_label.config(text=f'RX: {self._rx_count}'))

    if msg.arb_id in self.db.messages:
        # 闭包陷阱解决：使用默认参数捕获当前值
        self.after(0, lambda mid=msg.arb_id, d=msg.data:
                   self.recv_panel.update_message(mid, d))
```

**闭包陷阱**：lambda中使用`mid=msg.arb_id, d=msg.data`作为默认参数，确保在创建时捕获值而非引用。

### 5.9 信号值更新（闪烁效果）

```python
def update_value(self, phys_val, raw):
    # 设置文本和蓝色高亮
    self.val_lbl.config(text=text, foreground='#0066CC')
    # 200ms后恢复默认颜色
    if self._flash_job:
        self.after_cancel(self._flash_job)
    self._flash_job = self.after(200, lambda: self.val_lbl.config(
        foreground='#333333') if self.winfo_exists() else None)
```

### 5.10 SendPanel Category分组

```python
CATEGORIES = [
    ('Seat Control - Front', '#1B5E20', [710, 711]),   # CCU_FSCMCTRL, CCU_RSCMCTRL
    ('Mirror Control', '#1565C0', [712]),                # CCU_MirrorCTRL
    ('Vehicle Status', '#6A1B9A', [1412, 1264, 1265, 1296, 545, 291]),
    ('Body & Safety', '#E65100', [784, 801, 848, 1448, 706]),
]
# 未匹配的消息归入 'Other Signals'
```

### 5.11 循环发送机制

```python
# CanManager.start_cyclic_send 使用 threading.Timer 递归调度
def start_cyclic_send(self, arb_id, data_func, interval_ms):
    def _cyclic():
        if not self.is_connected:
            return
        data = data_func()
        if data:
            self.send_message(arb_id, data)
        # 递归调度下一次
        timer = threading.Timer(interval_ms / 1000.0, _cyclic)
        timer.daemon = True
        timer.start()
        self._send_timers[arb_id] = timer

    timer = threading.Timer(0, _cyclic)  # 立即执行第一次
    timer.daemon = True
    timer.start()
```

---

## 6. Excel → DBC 生成器（generate_dbc_from_excel.py）

### 6.1 Excel通信矩阵列定义

| 列号 | 列名 | 说明 |
|------|------|------|
| 1 | Msg Name | 消息名称 |
| 2 | Msg Type | CANFD Standard等 |
| 3 | Msg ID | 十六进制如 `0x2C6` |
| 4 | Msg Send Type | Cycle/Event/IfActive/CE/CA |
| 5 | Msg Cycle Time (ms) | 周期(ms) |
| 6 | Msg Length (Byte) | 数据长度 |
| 7 | Signal Name | 信号名称 |
| 8 | Multiplexing/Value | 多路复用 |
| 13 | Signal Description | 信号描述/注释 |
| 14 | **Byte Order** | Intel/Motorola MSB |
| 16 | **Start Bit** | 起始位（flat bit position） |
| 18 | **Signal Length (Bit)** | 信号长度 |
| 19 | Date Type | unsigned/signed |
| 20 | **Resolution** | 分辨率（factor） |
| 21 | **Offset** | 偏移量 |
| 26 | **Initial Value (Hex)** | 初始值（0x格式） |
| 27 | Invalid Value (Hex) | 无效值 |
| 29 | Unit | 单位 |
| 30 | Signal Value Description | 枚举值描述 `0x0:Text\r\n0x1:Text` |
| 34+ | Node columns | FSCM/CCU等，值为Tx/Rx |

### 6.2 值描述解析

```python
# Excel格式: "0x0:P\r\n0x1:L_Reserved\r\n0x2:2_Reserved\r\n..."
def parse_val_desc(text):
    result = {}
    parts = re.split(r'[\r\n\n]+', text)
    for part in parts:
        m = re.match(r'(0x[0-9A-Fa-f]+)\s*[:=]?\s*(.*)', part)
        if m:
            val = int(m.group(1), 16)
            desc = m.group(2).strip()
            result[val] = desc
    return result
```

### 6.3 发送节点确定

```python
# 扫描第34列之后的节点列，找第一个"Tx"
for col in range(34, ws.max_row + 1):
    cell_val = ws.cell(row=row_idx, column=col).value
    if cell_val and str(cell_val).strip() == 'Tx':
        node_name = ws.cell(row=1, column=col).value
        transmitter = str(node_name).strip()
        break
```

### 6.4 NM消息检测

```python
is_nm = '_NM_' in msg_name.upper() or msg_name.upper().endswith('_NM')
```

### 6.5 DBC名称约束

- 消息名称：最多32字符，只含`[a-zA-Z0-9_]`
- 信号名称：最多32字符，同上
- 非字母开头的加前缀`M_`或`S_`

### 6.6 运行命令

```bash
py "panel/CAN_TestTool/generate_dbc_from_excel.py"
```

---

## 7. 性能与注意事项

### 7.1 关键陷阱
1. **字节序**：DBC中byte_order=0是Intel (Little Endian)，=1是Motorola (Big Endian)。几乎所有车用ECU使用Motorola MSB
2. **闭包陷阱**：在RX回调中`self.after(0, lambda: func(x))` → `x`在lambda执行时才求值。必须用`lambda x=x: func(x)`捕获当前值
3. **线程安全**：tkinter不是线程安全的，所有UI操作必须通过`after()`在主线程执行
4. **GBK编码**：Windows中文系统DBC文件和Excel读取需用`encoding='gbk'`
5. **Emoji/Unicode**：GBK控制台不兼容`✓✗😊`等字符，print时需注意
6. **RX回调注册时序**：`register_rx_all_callback`必须在`connect_*()`之前调用？No — 修复方案是pending回调机制，无论先后均可

### 7.2 性能优化
- 信号值不变时不更新UI（`_last_values`缓存）
- RX线程中使用`timeout=0.05`避免忙等
- 闪烁效果使用`after_cancel`防止队列堆积

### 7.3 python-can 后端对照

| 后端 | bustype | 硬件 |
|------|---------|------|
| PCAN | `'pcan'` | PEAK PCAN-USB |
| Vector | `'vector'` | VN1610等 |
| SocketCAN | `'socketcan'` | Linux only |
| Virtual | `'virtual'` | 软件虚拟 |

### 7.4 常用DBC消息ID（FSCM项目）

| ID(hex) | ID(dec) | 名称 | 方向 |
|---------|---------|------|------|
| 0x123 | 291 | — | to_fscm |
| 0x221 | 545 | — | to_fscm |
| 0x2C2 | 706 | — | to_fscm |
| 0x310 | 784 | — | to_fscm |
| 0x321 | 801 | — | to_fscm |
| 0x350 | 848 | — | to_fscm |
| 0x50F | 1295 | FSCM_Right_0x50F | from_fscm (接收) |
| 0x511 | 1297 | FSCM_Mirror0x511 | from_fscm (接收) |
| 0x584 | 1412 | — | to_fscm |
| 0x5A0 | 1440 | — | to_fscm |
| 0x5A8 | 1448 | — | to_fscm |
| 0x2C6 | 710 | CCU_FSCMCTRL | to_fscm |
| 0x2C7 | 711 | CCU_RSCMCTRL | to_fscm |
| 0x2C8 | 712 | CCU_MirrorCTRL | to_fscm |
| 0x4F0 | 1264 | — | to_fscm |
| 0x4F1 | 1265 | — | to_fscm |

---

## 8. 快速模板

### 8.1 新建同类项目步骤
1. 准备Excel通信矩阵（参考第6节列定义）
2. 修改 `generate_dbc_from_excel.py` 中的 `EXCEL_PATH` 和 `OUTPUT_DBC`
3. 运行生成DBC
4. 修改 `dbc_parser.py` 中的 `classify_messages()` 调整消息分类规则
5. 修改 `can_test_tool.py` 中的 `SendPanel.CATEGORIES` 调整UI分组
6. 修改窗口标题和尺寸

### 8.2 最小信号解/编测试

```python
from dbc_parser import Signal

sig = Signal(
    name='TestSig', message_id=0x100,
    start_bit=7, size=8, byte_order=1,  # Motorola
    signed=False, factor=0.5, offset=0,
    minimum=0, maximum=100, unit='percent',
    receivers=['ECU']
)

# Encode
data = bytearray(8)
sig.pack(50.0, data)  # 50% → raw=100 → byte[0]=0x64

# Decode
value = sig.decode(data)  # 0x64 * 0.5 = 50.0
```

### 8.3 最小CAN连接模板

```python
from can_comms import CanManager

mgr = CanManager()
mgr.connect_virtual()  # 或 connect_pcan('PCAN_USBBUS1')

def on_rx(msg):
    print(f'RX: 0x{msg.arb_id:X} data={msg.data.hex()}')

mgr.register_rx_all_callback(on_rx)
mgr.send_message(0x100, b'\x01\x02\x03\x04')
mgr.disconnect()
```

### 8.4 最小GUI启动

```python
from dbc_parser import DbcParser
from can_test_tool import load_dbc, CanTestTool

db = load_dbc()  # 自动加载默认DBC
app = CanTestTool(db)
app.mainloop()
```
