"""
DBC (CAN Database) Parser
Parses DBC files into structured Python objects for the CAN test tool.
"""

import re
from dataclasses import dataclass, field
from typing import List, Dict, Optional


SEND_TYPE_NAMES = {
    0: 'Cycle', 1: 'NoMsgSendType', 2: 'IfActive',
    3: 'Event', 4: 'CA', 5: 'CE'
}


@dataclass
class Signal:
    """Represents a single CAN signal."""
    name: str
    message_id: int
    start_bit: int
    size: int
    byte_order: int  # 0=Intel(LittleEndian), 1=Motorola(BigEndian)
    signed: bool
    factor: float
    offset: float
    minimum: float
    maximum: float
    unit: str
    receivers: List[str]
    comment: str = ""
    value_descriptions: Dict[int, str] = field(default_factory=dict)
    gen_sig_start_value: Optional[float] = None
    gen_sig_invalid_value: Optional[int] = None
    gen_sig_send_type: int = 0

    def decode(self, data: bytes) -> float:
        """Decode raw CAN data bytes to physical value."""
        raw = self._extract_raw(data)
        return raw * self.factor + self.offset

    def encode(self, physical_value: float) -> int:
        """Encode physical value to raw integer."""
        raw = int((physical_value - self.offset) / self.factor + 0.5)
        max_raw = (1 << self.size) - 1
        return max(0, min(raw, max_raw))

    def _extract_raw(self, data: bytes) -> int:
        """Extract raw integer value from CAN data bytes."""
        if self.byte_order == 0:  # Intel (Little-Endian)
            return self._extract_intel(data)
        else:  # Motorola (Big-Endian)
            return self._extract_motorola(data)

    def _extract_intel(self, data: bytes) -> int:
        """Extract raw value using Intel (little-endian) bit ordering."""
        raw = 0
        for i in range(self.size):
            bit_pos = self.start_bit + i
            byte_idx = bit_pos // 8
            if byte_idx >= len(data):
                break
            bit_in_byte = bit_pos % 8
            if data[byte_idx] & (1 << bit_in_byte):
                raw |= (1 << i)
        return raw

    def _extract_motorola(self, data: bytes) -> int:
        """Extract raw value using Motorola (big-endian) bit ordering."""
        raw = 0
        for i in range(self.size):
            # Motorola: start_bit is MSB, bits go downward within byte, then to next lower byte
            msb_bit = self.start_bit
            bit_pos = msb_bit - i
            byte_idx = bit_pos // 8
            if byte_idx >= len(data):
                break
            bit_in_byte = 7 - (bit_pos % 8)
            if data[byte_idx] & (1 << bit_in_byte):
                raw |= (1 << i)
        return raw

    def pack(self, physical_value: float, data: bytearray):
        """Pack a physical value into CAN data bytes at the signal's position."""
        raw = self.encode(physical_value)
        max_val = (1 << self.size) - 1
        raw = max(0, min(raw, max_val))

        if self.byte_order == 0:  # Intel
            for i in range(self.size):
                bit_pos = self.start_bit + i
                byte_idx = bit_pos // 8
                if byte_idx >= len(data):
                    continue
                bit_in_byte = bit_pos % 8
                if raw & (1 << i):
                    data[byte_idx] |= (1 << bit_in_byte)
                else:
                    data[byte_idx] &= ~(1 << bit_in_byte)
        else:  # Motorola
            for i in range(self.size):
                msb_bit = self.start_bit
                bit_pos = msb_bit - i
                byte_idx = bit_pos // 8
                if byte_idx >= len(data):
                    continue
                bit_in_byte = 7 - (bit_pos % 8)
                if raw & (1 << i):
                    data[byte_idx] |= (1 << bit_in_byte)
                else:
                    data[byte_idx] &= ~(1 << bit_in_byte)

    def get_text_value(self, physical_value: float) -> str:
        """Get text description for enum values, or formatted physical value."""
        raw = self.encode(physical_value)
        if raw in self.value_descriptions:
            return f"{self.value_descriptions[raw]} ({physical_value})"
        if self.factor == 1 and self.offset == 0:
            return str(int(physical_value))
        return f"{physical_value:.1f}"


@dataclass
class Message:
    """Represents a single CAN message (frame)."""
    id: int
    name: str
    dlc: int
    transmitter: str
    signals: Dict[str, Signal] = field(default_factory=dict)
    cycle_time: int = 0
    send_type: int = 0
    vframe_format: int = 0
    is_nm: bool = False
    is_diag_request: bool = False
    is_diag_response: bool = False

    @property
    def send_type_name(self) -> str:
        return SEND_TYPE_NAMES.get(self.send_type, f'Type{self.send_type}')

    def get_signal_order(self) -> List[str]:
        """Get signal names ordered by start_bit (for consistent display)."""
        return sorted(self.signals.keys(), key=lambda n: self.signals[n].start_bit)


@dataclass
class DbcDatabase:
    """Complete parsed DBC database."""
    version: str = ""
    nodes: List[str] = field(default_factory=list)
    messages: Dict[int, Message] = field(default_factory=dict)

    def get_message_by_id(self, msg_id: int) -> Optional[Message]:
        return self.messages.get(msg_id)

    def get_message_by_name(self, name: str) -> Optional[Message]:
        for msg in self.messages.values():
            if msg.name == name:
                return msg
        return None


class DbcParser:
    """Parses DBC format text into DbcDatabase object."""

    # Regex patterns
    RE_VERSION = re.compile(r'VERSION\s+"([^"]*)"')
    RE_NODES = re.compile(r'BU_:\s*(.*)')
    RE_BO = re.compile(
        r'BO_\s+(\d+)\s+(\w+)\s*:\s*(\d+)\s+(\w+)'
    )
    RE_SG = re.compile(
        r'SG_\s+(\w+)\s*(?:\w+)?\s*:\s*(\d+)\|(\d+)@(\d)([+-])\s*\(([^)]+)\)\s*\[([^|]*)\|([^\]]*)\]\s*"([^"]*)"\s*(.*)'
    )
    RE_CM_SG = re.compile(
        r'CM_\s+SG_\s+(\d+)\s+(\w+)\s+"((?:[^"]|"")*)"\s*;'
    )
    RE_VAL = re.compile(
        r'VAL_\s+(\d+)\s+(\w+)\s+(.*?)\s*;'
    )
    RE_BA_CYCLE = re.compile(
        r'BA_\s+"GenMsgCycleTime"\s+BO_\s+(\d+)\s+(\d+)'
    )
    RE_BA_SENDTYPE = re.compile(
        r'BA_\s+"GenMsgSendType"\s+BO_\s+(\d+)\s+(\d+)'
    )
    RE_BA_VFRAME = re.compile(
        r'BA_\s+"VFrameFormat"\s+BO_\s+(\d+)\s+(\d+)'
    )
    RE_BA_NM = re.compile(
        r'BA_\s+"NmMessage"\s+BO_\s+(\d+)\s+(\d+)'
    )
    RE_BA_SIG_STARTVAL = re.compile(
        r'BA_\s+"GenSigStartValue"\s+SG_\s+(\d+)\s+(\w+)\s+(-?\d+)'
    )
    RE_BA_SIG_INVALID = re.compile(
        r'BA_\s+"GenSigInvalidValue"\s+SG_\s+(\d+)\s+(\w+)\s+(-?\d+)'
    )
    RE_BA_SIG_SENDTYPE = re.compile(
        r'BA_\s+"GenSigSendType"\s+SG_\s+(\d+)\s+(\w+)\s+(-?\d+)'
    )

    def parse(self, dbc_text: str) -> DbcDatabase:
        db = DbcDatabase()

        # Parse version
        m = self.RE_VERSION.search(dbc_text)
        if m:
            db.version = m.group(1)

        # Parse nodes
        m = self.RE_NODES.search(dbc_text)
        if m:
            db.nodes = m.group(1).strip().split()

        # Parse messages
        current_msg = None
        in_message = False

        for line in dbc_text.split('\n'):
            line_stripped = line.strip()

            # Parse BO_ (message definition)
            m = self.RE_BO.match(line_stripped)
            if m:
                msg_id = int(m.group(1))
                msg_name = m.group(2)
                dlc = int(m.group(3))
                transmitter = m.group(4)
                current_msg = Message(
                    id=msg_id,
                    name=msg_name,
                    dlc=dlc,
                    transmitter=transmitter
                )
                db.messages[msg_id] = current_msg
                in_message = True
                continue

            # Parse SG_ (signal definition)
            m = self.RE_SG.match(line_stripped)
            if m and current_msg is not None:
                sig_name = m.group(1)
                start_bit = int(m.group(2))
                size = int(m.group(3))
                byte_order = int(m.group(4))
                signed = m.group(5) == '-'

                # Parse factor,offset
                factor_str = m.group(6)
                if ',' in factor_str:
                    parts = factor_str.split(',')
                    factor = float(parts[0].strip())
                    offset = float(parts[1].strip())
                else:
                    factor = float(factor_str.strip())
                    offset = 0

                minimum = float(m.group(7).strip() or '0')
                maximum = float(m.group(8).strip() or '0')
                unit = m.group(9)
                receivers_str = m.group(10).strip()
                receivers = receivers_str.split(',') if receivers_str else []

                sig = Signal(
                    name=sig_name,
                    message_id=current_msg.id,
                    start_bit=start_bit,
                    size=size,
                    byte_order=byte_order,
                    signed=signed,
                    factor=factor,
                    offset=offset,
                    minimum=minimum,
                    maximum=maximum,
                    unit=unit,
                    receivers=receivers
                )
                current_msg.signals[sig_name] = sig
                continue

            # Check if we left the current message
            if line_stripped.startswith('BO_') and not m:
                in_message = False
                current_msg = None
            elif line_stripped.startswith('BU_'):
                in_message = False
                current_msg = None

        # Parse comments (CM_ SG_)
        for m in self.RE_CM_SG.finditer(dbc_text):
            msg_id = int(m.group(1))
            sig_name = m.group(2)
            comment = m.group(3).replace('""', '"')
            msg = db.messages.get(msg_id)
            if msg and sig_name in msg.signals:
                msg.signals[sig_name].comment = comment

        # Parse value descriptions (VAL_)
        for m in self.RE_VAL.finditer(dbc_text):
            msg_id = int(m.group(1))
            sig_name = m.group(2)
            val_str = m.group(3).strip()
            msg = db.messages.get(msg_id)
            if not msg or sig_name not in msg.signals:
                continue

            sig = msg.signals[sig_name]
            # Parse: value "text" value "text" ...
            pattern = re.compile(r'(\d+)\s+"((?:[^"]|"")*)"')
            for vm in pattern.finditer(val_str):
                val = int(vm.group(1))
                text = vm.group(2).replace('""', '"')
                sig.value_descriptions[val] = text

        # Parse attributes
        for m in self.RE_BA_CYCLE.finditer(dbc_text):
            msg_id = int(m.group(1))
            val = int(m.group(2))
            msg = db.messages.get(msg_id)
            if msg:
                msg.cycle_time = val

        for m in self.RE_BA_SENDTYPE.finditer(dbc_text):
            msg_id = int(m.group(1))
            val = int(m.group(2))
            msg = db.messages.get(msg_id)
            if msg:
                msg.send_type = val

        for m in self.RE_BA_VFRAME.finditer(dbc_text):
            msg_id = int(m.group(1))
            val = int(m.group(2))
            msg = db.messages.get(msg_id)
            if msg:
                msg.vframe_format = val

        for m in self.RE_BA_NM.finditer(dbc_text):
            msg_id = int(m.group(1))
            msg = db.messages.get(msg_id)
            if msg:
                msg.is_nm = True

        # Parse signal-level attributes
        for m in self.RE_BA_SIG_STARTVAL.finditer(dbc_text):
            msg_id = int(m.group(1))
            sig_name = m.group(2)
            val = int(m.group(3))
            msg = db.messages.get(msg_id)
            if msg and sig_name in msg.signals:
                sig = msg.signals[sig_name]
                sig.gen_sig_start_value = float(val * sig.factor + sig.offset) if sig else None

        for m in self.RE_BA_SIG_INVALID.finditer(dbc_text):
            msg_id = int(m.group(1))
            sig_name = m.group(2)
            val = int(m.group(3))
            msg = db.messages.get(msg_id)
            if msg and sig_name in msg.signals:
                msg.signals[sig_name].gen_sig_invalid_value = val

        for m in self.RE_BA_SIG_SENDTYPE.finditer(dbc_text):
            msg_id = int(m.group(1))
            sig_name = m.group(2)
            val = int(m.group(3))
            msg = db.messages.get(msg_id)
            if msg and sig_name in msg.signals:
                msg.signals[sig_name].gen_sig_send_type = val

        return db


def classify_messages(db: DbcDatabase):
    """Classify messages into 'to FSCM' (send controls) and 'from FSCM' (display)."""
    to_fscm = {}
    from_fscm = {}

    fscm_nodes = {'FSCM'}

    for msg_id, msg in db.messages.items():
        # Skip NM and diagnostic messages for the main UI
        if msg.is_nm:
            continue

        # Check if FSCM is a receiver of any signal
        fscm_is_receiver = any(
            'FSCM' in sig.receivers
            for sig in msg.signals.values()
        )

        # Check if FSCM is the transmitter
        fscm_is_transmitter = msg.transmitter == 'FSCM'

        if fscm_is_transmitter:
            from_fscm[msg_id] = msg
        elif fscm_is_receiver:
            to_fscm[msg_id] = msg

    return to_fscm, from_fscm
