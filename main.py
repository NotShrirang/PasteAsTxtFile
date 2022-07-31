import keyboard
import paste

while True:
    keyboard.add_hotkey("Ctrl+V", callback = lambda :paste.run())
    if paste.FLAG == 1:
        break