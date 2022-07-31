from pythoncom import CoInitialize
import subprocess
from urllib.parse import unquote
import win32clipboard
from win32com import client
import os

global FLAG
FLAG = 0

def run():
    global FLAG
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    CoInitialize()
    shell = client.Dispatch("Shell.Application")
    if len(list(shell.Windows())) == 0:
        with open( os.path.join(os.path.expanduser('~'), 'Desktop', 'mytxt.txt'), 'w') as f:
            f.write(data)
        FLAG = 1
        return
    else:
        unquoted_path = list(shell.Windows())[-1].LocationURL[len("file:///"):]
        most_recently_opened_window_path = unquote(unquoted_path)
        if not os.path.isdir(most_recently_opened_window_path):
            path = "/mytxt.txt"
            subprocess.Popen(f'explorer /select, {most_recently_opened_window_path + "/mytxt.txt"}')
        path = most_recently_opened_window_path + "/mytxt.txt"
        with open(path, "a") as f:
            f.write(data)
        FLAG = 1
    return
