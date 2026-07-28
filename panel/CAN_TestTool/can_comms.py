"""
CAN Communication Interface
Supports PCAN (via python-can with pcan backend) and virtual/cansee modes for testing.
"""

import threading
import time
import logging
from typing import Callable, Optional, Dict, Any
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)


@dataclass
class CanMessage:
    """CAN message data container."""
    arb_id: int
    data: bytes
    dlc: int = 8
    is_extended: bool = False
    timestamp: float = 0.0


class CanInterface:
    """Abstract base for CAN interface."""

    def __init__(self):
        self._rx_callbacks: Dict[int, list] = {}  # arb_id -> [callbacks]
        self._rx_all_callbacks: list = []
        self._running = False
        self._rx_thread: Optional[threading.Thread] = None

    def open(self, channel: str = '0', bitrate: int = 500000) -> bool:
        raise NotImplementedError

    def close(self):
        raise NotImplementedError

    def send(self, msg: CanMessage) -> bool:
        raise NotImplementedError

    def is_open(self) -> bool:
        raise NotImplementedError

    @property
    def name(self) -> str:
        return "Abstract"

    def register_rx_callback(self, arb_id: int, callback: Callable[[CanMessage], None]):
        if arb_id not in self._rx_callbacks:
            self._rx_callbacks[arb_id] = []
        self._rx_callbacks[arb_id].append(callback)

    def register_rx_all_callback(self, callback: Callable[[CanMessage], None]):
        self._rx_all_callbacks.append(callback)

    def unregister_rx_callback(self, arb_id: int, callback: Callable[[CanMessage], None]):
        if arb_id in self._rx_callbacks and callback in self._rx_callbacks[arb_id]:
            self._rx_callbacks[arb_id].remove(callback)

    def _notify_rx(self, msg: CanMessage):
        for cb in self._rx_all_callbacks:
            try:
                cb(msg)
            except Exception as e:
                logger.error(f"RX all callback error: {e}")

        if msg.arb_id in self._rx_callbacks:
            for cb in self._rx_callbacks[msg.arb_id]:
                try:
                    cb(msg)
                except Exception as e:
                    logger.error(f"RX callback error for 0x{msg.arb_id:X}: {e}")


class PcanInterface(CanInterface):
    """PCAN interface using python-can library."""

    def __init__(self):
        super().__init__()
        self._bus = None
        self._channel_map = {
            'PCAN_USBBUS1': 'PCAN_USBBUS1',
            'PCAN_USBBUS2': 'PCAN_USBBUS2',
            'PCAN_ISABUS1': 'PCAN_ISABUS1',
            'PCAN_ISABUS2': 'PCAN_ISABUS2',
            'PCAN_ISABUS3': 'PCAN_ISABUS3',
            'PCAN_ISABUS4': 'PCAN_ISABUS4',
            'PCAN_ISABUS5': 'PCAN_ISABUS5',
            'PCAN_ISABUS6': 'PCAN_ISABUS6',
            'PCAN_ISABUS7': 'PCAN_ISABUS7',
            'PCAN_ISABUS8': 'PCAN_ISABUS8',
        }

    def open(self, channel: str = 'PCAN_USBBUS1', bitrate: int = 500000,
             fd_mode: bool = False, data_bitrate: int = 2000000) -> bool:
        try:
            import can
            # Support numeric channel (e.g. "82") or named channel
            if channel.isdigit():
                chan = channel
            else:
                chan = self._channel_map.get(channel, channel)
            kwargs = dict(bustype='pcan', channel=chan, bitrate=bitrate)
            if fd_mode:
                kwargs['fd'] = True
                kwargs['data_bitrate'] = data_bitrate
            self._bus = can.interface.Bus(**kwargs)
            self._running = True
            self._rx_thread = threading.Thread(target=self._rx_loop, daemon=True)
            self._rx_thread.start()
            mode = 'CAN FD' if fd_mode else 'CAN'
            logger.info(f"PCAN opened: {channel} @ {bitrate} bps {mode}")
            return True
        except Exception as e:
            logger.error(f"Failed to open PCAN: {e}")
            self._bus = None
            return False

    def close(self):
        self._running = False
        if self._rx_thread:
            self._rx_thread.join(timeout=1)
        if self._bus:
            try:
                self._bus.shutdown()
            except Exception as e:
                logger.error(f"Error closing PCAN: {e}")
            self._bus = None

    def send(self, msg: CanMessage) -> bool:
        if not self._bus:
            return False
        try:
            import can
            cmsg = can.Message(
                arbitration_id=msg.arb_id,
                data=msg.data[:msg.dlc],
                dlc=msg.dlc,
                is_extended_id=msg.is_extended
            )
            self._bus.send(cmsg)
            return True
        except Exception as e:
            logger.error(f"PCAN send error: {e}")
            return False

    def is_open(self) -> bool:
        return self._bus is not None

    @property
    def name(self) -> str:
        return "PCAN"

    def _rx_loop(self):
        while self._running:
            try:
                msg = self._bus.recv(timeout=0.05)
                if msg is not None:
                    cmsg = CanMessage(
                        arb_id=msg.arithmetic_id if hasattr(msg, 'arithmetic_id') else msg.arbitration_id,
                        data=bytes(msg.data) if msg.data else b'',
                        dlc=msg.dlc,
                        is_extended=msg.is_extended_id if hasattr(msg, 'is_extended_id') else False,
                        timestamp=msg.timestamp if hasattr(msg, 'timestamp') else time.time()
                    )
                    self._notify_rx(cmsg)
            except Exception as e:
                if self._running:
                    logger.error(f"PCAN rx loop error: {e}")
                    time.sleep(0.01)


class VirtualCanInterface(CanInterface):
    """
    Virtual CAN interface for testing without hardware.
    In a real scenario, this could be replaced with SocketCAN or a virtual bus.
    Uses a simple loopback mechanism.
    """

    def __init__(self):
        super().__init__()
        self._opened = False
        self._rx_buffer: list = []
        self._buffer_lock = threading.Lock()

    def open(self, channel: str = 'virtual', bitrate: int = 500000) -> bool:
        self._opened = True
        self._running = True
        self._rx_thread = threading.Thread(target=self._rx_loop, daemon=True)
        self._rx_thread.start()
        logger.info("Virtual CAN opened")
        return True

    def close(self):
        self._running = False
        if self._rx_thread:
            self._rx_thread.join(timeout=1)
        self._opened = False

    def send(self, msg: CanMessage) -> bool:
        logger.debug(f"[Virtual TX] 0x{msg.arb_id:X}: {' '.join(f'{b:02X}' for b in msg.data[:msg.dlc])}")
        # Loopback: queue the message for RX
        with self._buffer_lock:
            self._rx_buffer.append(msg)
        return True

    def is_open(self) -> bool:
        return self._opened

    @property
    def name(self) -> str:
        return "Virtual"

    def _rx_loop(self):
        while self._running:
            msgs = []
            with self._buffer_lock:
                msgs = list(self._rx_buffer)
                self._rx_buffer.clear()
            for msg in msgs:
                self._notify_rx(msg)
            time.sleep(0.01)


class CanoeInterface(CanInterface):
    """CAN interface using Vector hardware (VN1610, etc.) via python-can vector backend.
    Also compatible with CANoe software using same Vector hardware."""

    def __init__(self):
        super().__init__()
        self._bus = None

    def open(self, channel: str = '0', bitrate: int = 500000,
             fd_mode: bool = False, data_bitrate: int = 2000000) -> bool:
        try:
            import can
            kwargs = dict(
                bustype='vector',
                channel=channel,
                bitrate=bitrate,
                app_name='CAN Test Tool'
            )
            if fd_mode:
                kwargs['fd'] = True
                kwargs['data_bitrate'] = data_bitrate
            self._bus = can.interface.Bus(**kwargs)
            self._running = True
            self._rx_thread = threading.Thread(target=self._rx_loop, daemon=True)
            self._rx_thread.start()
            mode = 'CAN FD' if fd_mode else 'CAN'
            logger.info(f"CANoe/Vector opened: channel={channel} @ {bitrate} bps {mode}")
            return True
        except Exception as e:
            logger.error(f"Failed to open CANoe/Vector: {e}")
            self._bus = None
            return False

    def close(self):
        self._running = False
        if self._rx_thread:
            self._rx_thread.join(timeout=1)
        if self._bus:
            try:
                self._bus.shutdown()
            except Exception as e:
                logger.error(f"Error closing CANoe/Vector: {e}")
            self._bus = None

    def send(self, msg: CanMessage) -> bool:
        if not self._bus:
            return False
        try:
            import can
            cmsg = can.Message(
                arbitration_id=msg.arb_id,
                data=msg.data[:msg.dlc],
                dlc=msg.dlc,
                is_extended_id=msg.is_extended
            )
            self._bus.send(cmsg)
            return True
        except Exception as e:
            logger.error(f"CANoe/Vector send error: {e}")
            return False

    def is_open(self) -> bool:
        return self._bus is not None

    @property
    def name(self) -> str:
        return "CANoe/Vector"

    def _rx_loop(self):
        while self._running:
            try:
                msg = self._bus.recv(timeout=0.05)
                if msg is not None:
                    cmsg = CanMessage(
                        arb_id=msg.arbitration_id,
                        data=bytes(msg.data) if msg.data else b'',
                        dlc=msg.dlc,
                        is_extended=msg.is_extended_id if hasattr(msg, 'is_extended_id') else False,
                        timestamp=msg.timestamp if hasattr(msg, 'timestamp') else time.time()
                    )
                    self._notify_rx(cmsg)
            except Exception as e:
                if self._running:
                    logger.error(f"CANoe/Vector rx loop error: {e}")
                    time.sleep(0.01)


class CanManager:
    """
    Manages CAN communication lifecycle.
    Supports PCAN hardware and virtual mode.
    """

    @staticmethod
    def list_available_devices() -> list:
        """Scan for available CAN devices (PCAN + Vector) on the system.
        Returns list of channel strings with type hints.
        """
        devices = []
        try:
            import can
            # Scan PCAN
            configs = can.detect_available_configs(bustype='pcan')
            for cfg in configs:
                ch = cfg.get('channel', '')
                if ch:
                    if ch.startswith('pcan:'):
                        ch = ch[5:]
                    devices.append(ch)
            # Scan Vector
            configs = can.detect_available_configs(bustype='vector')
            for cfg in configs:
                ch = cfg.get('channel', '')
                if ch and ch not in devices:
                    devices.append(ch)
        except Exception:
            pass
        # Ensure standard PCAN channels are listed
        standard = ['PCAN_USBBUS1', 'PCAN_USBBUS2', 'PCAN_ISABUS1', 'PCAN_ISABUS2']
        for s in standard:
            if s not in devices:
                devices.append(s)
        devices.append('virtual')
        return devices

    def __init__(self):
        self._interface: Optional[CanInterface] = None
        self._send_timers: Dict[int, threading.Timer] = {}
        self._send_timer_lock = threading.Lock()
        self._signal_values: Dict[int, Dict[str, float]] = {}  # msg_id -> {sig_name: phys_val}
        # Pending RX callbacks - registered before interface is created
        self._pending_rx_all_callbacks: list = []
        self._pending_rx_callbacks: Dict[int, list] = {}

    def connect_pcan(self, channel: str = 'PCAN_USBBUS1', bitrate: int = 500000,
                      fd_mode: bool = False, data_bitrate: int = 2000000) -> bool:
        if self._interface and self._interface.is_open():
            self.disconnect()
        self._interface = PcanInterface()
        ok = self._interface.open(channel, bitrate, fd_mode, data_bitrate)
        if ok:
            self._flush_pending_rx_callbacks()
        return ok

    def connect_virtual(self) -> bool:
        if self._interface and self._interface.is_open():
            self.disconnect()
        self._interface = VirtualCanInterface()
        ok = self._interface.open()
        if ok:
            self._flush_pending_rx_callbacks()
        return ok

    def connect_canoe(self, channel: str = '0', bitrate: int = 500000,
                       fd_mode: bool = False, data_bitrate: int = 2000000) -> bool:
        """Connect using CANoe/Vector hardware."""
        if self._interface and self._interface.is_open():
            self.disconnect()
        self._interface = CanoeInterface()
        ok = self._interface.open(channel, bitrate, fd_mode, data_bitrate)
        if ok:
            self._flush_pending_rx_callbacks()
        return ok

    def disconnect(self):
        self.stop_all_cyclic()
        if self._interface:
            self._interface.close()
            self._interface = None

    @property
    def is_connected(self) -> bool:
        return self._interface is not None and self._interface.is_open()

    @property
    def interface_name(self) -> str:
        if self._interface:
            return self._interface.name
        return "None"

    def send_signal(self, msg_id: int, signal_name: str, value: float, dlc: int = 8):
        """Send a single signal (others keep their previous values or zero)."""
        if not self._interface:
            return False

        # Build the message data
        if msg_id not in self._signal_values:
            self._signal_values[msg_id] = {}

        self._signal_values[msg_id][signal_name] = value

        # Get the full message from dbc info (will be provided by the GUI layer)
        # This is a stub - the actual encoding happens at the GUI level
        return True

    def send_message(self, arb_id: int, data: bytes, dlc: int = 8, is_extended: bool = False) -> bool:
        if not self._interface:
            return False
        msg = CanMessage(arb_id=arb_id, data=data, dlc=dlc, is_extended=is_extended)
        return self._interface.send(msg)

    def register_rx_callback(self, arb_id: int, callback: Callable[[CanMessage], None]):
        """Register a per-ID RX callback, storing it for later interfaces if needed."""
        if arb_id not in self._pending_rx_callbacks:
            self._pending_rx_callbacks[arb_id] = []
        self._pending_rx_callbacks[arb_id].append(callback)
        if self._interface:
            self._interface.register_rx_callback(arb_id, callback)

    def register_rx_all_callback(self, callback: Callable[[CanMessage], None]):
        """Register an all-messages RX callback, storing for later interfaces if needed."""
        self._pending_rx_all_callbacks.append(callback)
        if self._interface:
            self._interface.register_rx_all_callback(callback)

    def _flush_pending_rx_callbacks(self):
        """Re-register all pending RX callbacks with the current interface."""
        if not self._interface:
            return
        for cb in self._pending_rx_all_callbacks:
            self._interface.register_rx_all_callback(cb)
        for arb_id, cbs in self._pending_rx_callbacks.items():
            for cb in cbs:
                self._interface.register_rx_callback(arb_id, cb)

    def start_cyclic_send(self, arb_id: int, data_func, interval_ms: int):
        """
        Start cyclic sending of a message.
        data_func: callable that returns bytes to send
        """
        self.stop_cyclic_send(arb_id)

        def _cyclic():
            if not self.is_connected:
                return
            try:
                data = data_func()
                self.send_message(arb_id, data)
            except Exception as e:
                logger.error(f"Cyclic send error 0x{arb_id:X}: {e}")
            with self._send_timer_lock:
                timer = threading.Timer(interval_ms / 1000.0, _cyclic)
                timer.daemon = True
                timer.start()
                self._send_timers[arb_id] = timer

        timer = threading.Timer(0, _cyclic)
        timer.daemon = True
        timer.start()
        with self._send_timer_lock:
            self._send_timers[arb_id] = timer

    def stop_cyclic_send(self, arb_id: int):
        with self._send_timer_lock:
            if arb_id in self._send_timers:
                self._send_timers[arb_id].cancel()
                del self._send_timers[arb_id]

    def stop_all_cyclic(self):
        with self._send_timer_lock:
            for timer in self._send_timers.values():
                timer.cancel()
            self._send_timers.clear()
