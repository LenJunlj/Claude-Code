# excel-to-ldf — 从 LIN 通信矩阵 Excel 生成 Vector LDF Explorer Pro 兼容的 LDF 文件

## 功能

解析 EEA 5.1 格式的 LIN 通信矩阵 Excel（`.xlsx`），生成 Vector LDF Explorer Pro 可直接打开的 `.ldf` 文件。

## 用法

```bash
python excel_to_ldf.py <输入Excel路径> [输出LDF路径]
```

如果不指定输出路径，默认在 Excel 同级目录生成同名 `.ldf` 文件。

## 依赖

```bash
pip install openpyxl
```

---

## 生成规则（踩坑记录）

以下规则来自实际调试 Vector LDF Explorer Pro 的过程，**缺一不可**。

### 1. 文件格式

- **编码**：UTF-8 **without BOM**（有 BOM 会导致 NullReferenceException）
- **行尾**：CRLF（`\r\n`）
- **开头**：两个空行（CRLF CRLF）后接 `LIN_description_file;`
- **结尾**：最后一个 `}` 后跟一个 CRLF
- **注释**：不写注释（工具解析时可能出问题）

### 2. 全局声明

```ldf
LIN_description_file;                          // 不带文件名参数！
LIN_protocol_version = "2.0";                  // 即使 Excel 写 2.1，也要声明 2.0
LIN_language_version = "2.0";
LIN_speed = 19.2 kbps;                         // 从 Excel Info sheet 读取
```

### 3. Nodes

```ldf
Nodes {
  Master: FLZCU, 5 ms, 0.1 ms ;               // 分号前有空格
  Slaves: SRF ;
}
```

- Master 节点的 NAD 在 Excel 中标记为 `-`
- Slave 节点的 NAD 从 Excel Info sheet 读取（如 `0x01`）

### 4. Signals

```ldf
Signals {
  SignalName: bit_length, init_value, publisher, subscriber ;
}
```

**关键规则**：
- 参数顺序：**length, init_value, publisher, subscriber**（不含 start_bit！）
- **init_value 必须用十进制**（`0` 而不是 `0x0`）
- start_bit 不在信号定义中指定，而是在 `Frames` 中指定
- publisher/subscriber 根据 Excel Matrix 中对应列的 `s`（发送）/ `r`（接收）确定

### 5. Diagnostic_signals

```ldf
Diagnostic_signals {
  MasterReqB0: 8, 0 ;
  ...
  SlaveRespB7: 8, 0 ;
}
```

固定 8 字节诊断信号，每个信号 8 位。

### 6. Frames

```ldf
Frames {
  FrameName: frame_id_decimal, publisher, length {
    SignalName, start_bit ;
  }
}
```

**关键规则**：
- **帧 ID 必须用十进制**（`48` 而不是 `0x30`）
- 空帧用 `{ }` 表示（大括号单独一行）
- start_bit 在帧的信号条目中指定

### 7. Diagnostic_frames

```ldf
Diagnostic_frames {
  MasterReq: 0x3c { ... }                      // 诊帧 ID 用十六进制
  SlaveResp: 0x3d { ... }
}
```

### 8. Node_attributes

```ldf
Node_attributes {
  SlaveName{                                   // 大括号在同一行，无空格
    LIN_protocol = "2.0" ;
    configured_NAD = 0x01 ;
    product_id = 0x0, 0x0, 255 ;
    P2_min = 0 ms ;                            // 必须包含
    ST_min = 0 ms ;                            // 必须包含
    configurable_frames {
      FrameName = 0x0 ;
    }
  }
}
```

### 9. Schedule_tables

```ldf
Schedule_tables {
 ScheduleName {                                // 缩进一个空格
    FrameName delay 10 ms ;
  }
}
```

延迟时间从 Excel LIN Schedule sheet 读取。

### 10. Signal_encoding_types（最容易出错的部分）

```ldf
Signal_encoding_types {
  SignalName_Encoding {                        // 编码名：信号名 + _Encoding 后缀
    logical_value, 0, "description" ;
    logical_value, 1, "description2" ;
  }
}
```

**关键规则**：

#### 10a. 有值描述（value_description）的信号
从 Excel value_description 列解析 `0x0:desc\n0x1:desc2...` 格式，生成 `logical_value` 条目。

#### 10b. 有物理值编码（resolution ≠ 1 或 offset ≠ 0）的信号
```ldf
  VehicleSpeedVSOSig_Encoding {
    physical_value, 0, 8190, 0.0625, 0 ;       // min_raw, max_raw, scale, offset
    logical_value, 0, "0.0" ;                   // 必须同时有 logical_value！
    logical_value, 8190, "511.875" ;
  }
```

**🚨 此处极易出错：**
- `physical_value` 的 min/max 必须是 **raw 值（整数）**，不是物理值（浮点数）❌
  - 正确：`physical_value, 0, 8190, 0.0625, 0 ;`（raw 值整数）
  - 错误：`physical_value, 0.0, 511.875, 0.0625, 0 ;`（物理值浮点数 → NullReferenceException）
- **physical_value 不能单独存在**，必须同时有 `logical_value` 条目
  - 只有 physical_value 没有 logical_value → NullReferenceException

#### 10c. 既无值描述又无特殊分辨率的信号（如 SoftwareVersion）
**不创建 encoding 条目**——在 Signal_representation 中也跳过。

### 11. Signal_representation

```ldf
Signal_representation {
  EncodingName: SignalName ;                   // 冒号 + 空格，编码名在左
}
```

**关键规则**：
- 格式：**`encoding_name : signal_name ;`**（冒号分隔，编码左信号右）
- ❌ 错误：`signal_name, encoding_name;`（逗号分隔且顺序反了）

### 12. 括号平衡

生成后务必验证：`{` 数量 = `}` 数量。

```
Brackets: { = 49, } = 49, balanced = True
```

---

## 排查方法论

如果生成的文件打不开（NullReferenceException），用二分法逐步排错：

1. **最小可工作版本**：1 信号 + 1 帧 + 1 encoding，确认基础结构正确
2. **逐步增加信号**：先加同一帧内更多信号
3. **逐步增加帧**：加第二帧、第三帧
4. **逐步增加 schedule**：加调度表条目
5. **逐步增加 encoding 类型**：先加 logical_value 类型，最后加 physical_value 类型
6. **检查 `physical_value` 参数**：必须用 raw 值（整数）+ 搭配 logical_value
