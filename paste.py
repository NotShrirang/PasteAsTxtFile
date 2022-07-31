from pythoncom import CoInitialize
from urllib.parse import unquote
import win32clipboard
from win32com import client

global FLAG
FLAG = 0

def run():
    global FLAG
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    CoInitialize()
    shell = client.Dispatch("Shell.Application")
    unquoted_path = list(shell.Windows())[-1].LocationURL[len("file:///"):]
    most_recently_opened_window_path = unquote(unquoted_path)
    path = most_recently_opened_window_path + "/mytxt.txt"
    with open(path, "a") as f:
        f.write(data)
    FLAG = 1
    return