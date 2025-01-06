import os
import psutil
import pyautogui
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import datetime as dt

class SystemManager:
    def __init__(self):
        self._initialize_audio()
        self._initialize_screenshot_dir()

    def _initialize_audio(self):
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(IAudioEndpointVolume._iid_, 1, None)
        self.volume = self.interface.QueryInterface(IAudioEndpointVolume)

    def _initialize_screenshot_dir(self):
        self.screenshots_dir = os.path.join(os.path.expanduser("~"), "Pictures", "Luna_Screenshots")
        if not os.path.exists(self.screenshots_dir):
            os.makedirs(self.screenshots_dir)

    def control_system_volume(self, action):
        actions = {
            "up": lambda: self.volume.SetMasterVolumeLevelScalar(
                min(self.volume.GetMasterVolumeLevelScalar() + 0.1, 1.0), None),
            "down": lambda: self.volume.SetMasterVolumeLevelScalar(
                max(self.volume.GetMasterVolumeLevelScalar() - 0.1, 0.0), None),
            "mute": lambda: self.volume.SetMute(1, None),
            "unmute": lambda: self.volume.SetMute(0, None)
        }
        if action in actions:
            actions[action]()
        else:
            print("Unknown action:", action)

    def get_system_info(self):
        try:
            cpu = psutil.cpu_percent(interval=1)
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            battery = psutil.sensors_battery()

            info = (f"CPU usage: {cpu}%. "
                   f"Memory usage: {memory.percent}%. "
                   f"Disk usage: {disk.percent}%.")

            if battery:
                info += (f" Battery at {battery.percent}% "
                        f"({'charging' if battery.power_plugged else 'not charging'})")

            return info
        except Exception as e:
            return f"Error getting system information: {str(e)}"

    def take_screenshot(self):
        try:
            timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshot_{timestamp}.png"
            filepath = os.path.join(self.screenshots_dir, filename)
            screenshot = pyautogui.screenshot()
            screenshot.save(filepath)
            return filepath
        except Exception as e:
            print(f"Error taking screenshot: {e}")
            return None
