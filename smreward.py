import psutil
import os, subprocess, time, toml;
from windows import WindowsHelper


# Config settings
START_UP_TIME = 300
UPDATE_ATTEMPTS = 60
SHUTDOWN_AFTER = True


def get_steam_exe_path():
    # try known install locations for all Disk Drives
    installLocations = ["X:\\Steam\\steam.exe", "X:\\Program Files\\Steam\\steam.exe", "X:\\Program Files (x86)\\Steam\\steam.exe"]
    for drive in WindowsHelper.getDiskDrives():
        for path in installLocations:
            path = path.replace("X", drive)
            if os.path.exists(path):
                return path

    # check registry for install location
    reg1 = WindowsHelper.readRegistry(p = "SOFTWARE\\Wow6432Node\\Valve\\Steam", k = 'InstallPath')
    reg2 = WindowsHelper.readRegistry(p = "SOFTWARE\\Valve\\Steam", k = 'InstallPath')
    possiblePaths = [os.path.join(i, 'steam.exe') for i in [reg1, reg2] if i is not None]
    for path in possiblePaths:
        if os.path.exists(path):
            return path

    # try to find it via the Start Menu shortcut
    startmenuPath = "{}\\Microsoft\\Windows\\Start Menu\\Programs\\Steam\\Steam.lnk".format(os.getenv('APPDATA'))
    if os.path.exists(startmenuPath):
        path = WindowsHelper.followShortcut(startmenuPath)
        if os.path.exists(path):
            return path
    
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

def fore_ground_smite():
    pid = get_smite()
    handles = WindowsHelper.getHandlesForProcess(pid)

    for hwnd in handles:
        if WindowsHelper.IsWindowVisible(hwnd):
            WindowsHelper.SetWindowForeground(hwnd)

def smite_in_focus():
    pid = get_smite()
    return WindowsHelper.processInFocus(pid)


# Main function that brings every

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
    time.sleep(60)

    print("Collecting reward")
    count = 0
    while count < 20:
        WindowsHelper.InputKey(0x1B) # Escape Key
        time.sleep(2)
        count += 1

    if SHUTDOWN_AFTER:
        print("Shutting down")
        os.system("shutdown -s")
    else:
        print("Done")