"""
Dispatcher de plataforma — importa a implementação correta para Mac ou Windows.
  Mac     → word_bridge_mac.py  (AppleScript via osascript)
  Windows → word_bridge_win.py  (COM automation via pywin32)
"""
import sys

if sys.platform == "darwin":
    from app.word_bridge_mac import *          # noqa: F401, F403
elif sys.platform == "win32":
    from app.word_bridge_win import *          # noqa: F401, F403
else:
    raise RuntimeError(f"Plataforma não suportada: {sys.platform}")
