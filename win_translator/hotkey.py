import os
from dataclasses import dataclass

from PySide6 import QtCore


if os.name != 'nt':
    # Заглушка, щоб проєкт хоча б імпортувався не на Windows.
    @dataclass
    class HotkeySpec:
        ctrl: bool
        shift: bool
        alt: bool
        vk: int

    class GlobalHotkey(QtCore.QObject):
        triggered = QtCore.Signal()

        def __init__(self, hotkey_id: int, spec: HotkeySpec):
            super().__init__()

        def start(self) -> None:
            raise RuntimeError('GlobalHotkey працює лише на Windows')

        def stop(self) -> None:
            return

else:
    import ctypes
    import ctypes.wintypes

    user32 = ctypes.windll.user32

    WM_HOTKEY = 0x0312

    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_NOREPEAT = 0x4000


    @dataclass
    class HotkeySpec:
        ctrl: bool
        shift: bool
        alt: bool
        vk: int

        def modifiers(self) -> int:
            mods = 0
            if self.ctrl:
                mods |= MOD_CONTROL
            if self.shift:
                mods |= MOD_SHIFT
            if self.alt:
                mods |= MOD_ALT
            # Щоб не спамило при утриманні
            mods |= MOD_NOREPEAT
            return mods


    class _NativeHotkeyFilter(QtCore.QAbstractNativeEventFilter):
        def __init__(self, hotkey_id: int, signal: QtCore.SignalInstance):
            super().__init__()
            self._id = hotkey_id
            self._signal = signal

        def nativeEventFilter(self, eventType, message):
            # PySide6 на Windows дає MSG* в message (int)
            try:
                if eventType != 'windows_generic_MSG' and eventType != 'windows_dispatcher_MSG':
                    return False, 0

                msg_ptr = ctypes.cast(int(message), ctypes.POINTER(ctypes.wintypes.MSG))
                msg = msg_ptr.contents

                if msg.message == WM_HOTKEY and int(msg.wParam) == int(self._id):
                    self._signal.emit()
                    return True, 0
            except Exception:
                pass

            return False, 0


    class GlobalHotkey(QtCore.QObject):
        triggered = QtCore.Signal()

        def __init__(self, hotkey_id: int, spec: HotkeySpec):
            super().__init__()
            self._id = int(hotkey_id)
            self._spec = spec
            self._filter = None
            self._registered = False

        def start(self) -> None:
            if self._registered:
                return

            mods = int(self._spec.modifiers())
            vk = int(self._spec.vk)

            ok = user32.RegisterHotKey(None, self._id, mods, vk)
            if not ok:
                raise RuntimeError('Не вдалося зареєструвати глобальну гарячу клавішу (можливо, вже зайнята)')

            app = QtCore.QCoreApplication.instance()
            if app is None:
                # Теоретично не має статись, але Windows любить сюрпризи
                user32.UnregisterHotKey(None, self._id)
                raise RuntimeError('Qt application не ініціалізовано')

            self._filter = _NativeHotkeyFilter(self._id, self.triggered)
            app.installNativeEventFilter(self._filter)
            self._registered = True

        def stop(self) -> None:
            if not self._registered:
                return

            try:
                user32.UnregisterHotKey(None, self._id)
            except Exception:
                pass

            app = QtCore.QCoreApplication.instance()
            if app is not None and self._filter is not None:
                try:
                    app.removeNativeEventFilter(self._filter)
                except Exception:
                    pass

            self._filter = None
            self._registered = False
