import psutil
import win32gui, win32process, win32com.client, win32con
import subprocess, time;
import ctypes
import ctypes.wintypes as wintypes

IsWindowVisible = ctypes.windll.user32.IsWindowVisible

def start_smite():
    args = ["C:\Program Files (x86)\Steam\Steam.exe", "steam://rungameid/386360"]
    subprocess.Popen(args)


def smite_running():
    for p in psutil.process_iter():
        if p.name().lower() == "smite.exe":
            return True
    return False
        
def get_smite():
    for p in psutil.process_iter():
        if p.name().lower() == "smite.exe":
            return p.pid
    return None

def get_hwnds_for_pid(pid):
    def callback(hwnd, hwnds):
        #if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)

        if found_pid == pid:
            hwnds.append(hwnd)
        return True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds 


def fore_ground_smite():
    pid = get_smite()
    handles = get_hwnds_for_pid(pid)

    for hwnd in handles:
        if IsWindowVisible(hwnd):
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(hwnd)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) # needed for fullscreen
            return

def smite_in_focus():
    handle = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(handle)

    return pid == get_smite()

user32 = ctypes.WinDLL('user32', use_last_error=True)

INPUT_MOUSE    = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_UNICODE     = 0x0004
KEYEVENTF_SCANCODE    = 0x0008

MAPVK_VK_TO_VSC = 0

# msdn.microsoft.com/en-us/library/dd375731
VK_TAB  = 0x09
VK_MENU = 0x12
VK_ESCAPE = 0x1B

# C struct definitions

wintypes.ULONG_PTR = wintypes.WPARAM

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

    def __init__(self, *args, **kwds):
        super(KEYBDINPUT, self).__init__(*args, **kwds)
        # some programs use the scan code even if KEYEVENTF_SCANCODE
        # isn't set in dwFflags, so attempt to map the correct code.
        if not self.dwFlags & KEYEVENTF_UNICODE:
            self.wScan = user32.MapVirtualKeyExW(self.wVk,
                                                 MAPVK_VK_TO_VSC, 0)

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

LPINPUT = ctypes.POINTER(INPUT)

def _check_count(result, func, args):
    if result == 0:
        raise ctypes.WinError(ctypes.get_last_error())
    return args

user32.SendInput.errcheck = _check_count
user32.SendInput.argtypes = (wintypes.UINT, # nInputs
                             LPINPUT,       # pInputs
                             ctypes.c_int)  # cbSize

def PressKey(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode,
                            dwFlags=KEYEVENTF_KEYUP))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))


if __name__ == '__main__':

    print("============================")
    print("Running Smite Startup script")
    print("============================\n\n")

    print("Waiting for PC to start up.")
    time.sleep(60)


    if not smite_running():
        print("Starting smite.")
        start_smite()

    # 60 attempts is a half hour (should be enough for updates)
    attempts = 0
    while not smite_running() and attempts < 60:
        print("Waiting for smite...")
        time.sleep(30)
        attempts += 1

    print("Smite is running.")
    time.sleep(60)

    while not smite_in_focus():
        print("Smite not in focus, setting focus.")
        fore_ground_smite()
        time.sleep(20)

    # smite should be running and in focus
    print("Waiting for game to load.")
    time.sleep(60)

    print("Collecting reward")
    count = 0
    while count < 20:
        PressKey(win32con.VK_ESCAPE)
        ReleaseKey(win32con.VK_ESCAPE)
        time.sleep(2)
        count += 1

    print("Done")

