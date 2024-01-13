import psutil
import win32api, win32gui, win32process, win32com.client, win32con, winreg
import os, subprocess, time, toml;
import ctypes
import ctypes.wintypes as wintypes

IsWindowVisible = ctypes.windll.user32.IsWindowVisible

# Config settings
START_UP_TIME = 300
UPDATE_ATTEMPTS = 60
SHUTDOWN_AFTER = True

def read_reg(ep, p = "", k = ""):
    try:
        key = winreg.OpenKeyEx(ep, p)
        value = winreg.QueryValueEx(key,k)
        if key:
            winreg.CloseKey(key)
        return str(value[0])
    except:
        return None


def get_steam_exe_path():

    # try known install locations for all Disk Drives
    installLocations = ["X:\\Steam\\steam.exe", "X:\\Program Files\\Steam\\steam.exe", "X:\\Program Files (x86)\\Steam\\steam.exe"]
    for drive in win32api.GetLogicalDriveStrings().split('\000')[:-1]:
        drive = drive.replace(':\\', '')
        for path in installLocations:
            path = path.replace("X", drive)
            if os.path.exists(path):
                return path

    # check registry for install location
    reg1 = read_reg(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Valve\Steam", 'InstallPath')
    reg2 = read_reg(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Valve\Steam", 'InstallPath')
    possiblePaths = [os.path.join(i, 'steam.exe') for i in [reg1, reg2] if i is not None]
    for path in possiblePaths:
        if os.path.exists(path):
            return path

    # try to find it via the Start Menu shortcut
    startmenuPath = "{}\\Microsoft\\Windows\\Start Menu\\Programs\\Steam\\Steam.lnk".format(os.getenv('APPDATA'))
    if os.path.exists(startmenuPath):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(startmenuPath)
        if os.path.exists(shortcut.Targetpath):
            return shortcut.Targetpath
    
    # No path found
    return None

def start_smite():
    path = get_steam_exe_path()

    if path is None:
        print("Cannot find steam executable")
        exit()

    args = [path, "steam://rungameid/386360"]
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

    print("Reading configuration")
    if not os.path.exists('./config/settings.toml'):
        os.mkdir('config')
        with open('./config/settings.toml', 'w') as f:
            toml.dump({
                "program": {
                    "START_UP_TIME": 300,
                    "UPDATE_ATTEMPTS": 60,
                    "SHUTDOWN_AFTER": True
                }
            }, f)
    else:
        try:
            with open('./config/settings.toml', 'r') as f:
                config = toml.load(f)
            
            START_UP_TIME = int(config['program']['START_UP_TIME'])
            UPDATE_ATTEMPTS = int(config['program']['UPDATE_ATTEMPTS'])
            SHUTDOWN_AFTER = bool(config['program']['SHUTDOWN_AFTER'])
        except:
            print("Couldn't read config values. Using defaults.")

    print("Waiting for PC to start up.")
    time.sleep(START_UP_TIME)


    if not smite_running():
        print("Starting smite.")
        start_smite()

    attempts = 0
    while not smite_running() and attempts < UPDATE_ATTEMPTS:
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
    time.sleep(120)

    print("Collecting reward")
    count = 0
    while count < 20:
        PressKey(win32con.VK_ESCAPE)
        ReleaseKey(win32con.VK_ESCAPE)
        time.sleep(2)
        count += 1

    if SHUTDOWN_AFTER:
        print("Shutting down")
        os.system("shutdown -s")
    else:
        print("Done")