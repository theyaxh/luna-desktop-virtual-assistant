import os
import subprocess
import webbrowser

class AppManager:
    def __init__(self):
        self._load_resource_paths()

    def _load_resource_paths(self):
        self.apps = {
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "paint": "mspaint.exe",
            "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
            "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            "vlc": r"C:\Program Files\VideoLAN\VLC\vlc.exe",
            "control panel": "control",
            "file explorer": "explorer.exe"
        }

        self.common_dirs = {
            "desktop": os.path.join(os.path.expanduser("~"), "Desktop"),
            "documents": os.path.join(os.path.expanduser("~"), "Documents"),
            "downloads": os.path.join(os.path.expanduser("~"), "Downloads"),
            "music": os.path.join(os.path.expanduser("~"), "Music"),
            "pictures": os.path.join(os.path.expanduser("~"), "Pictures"),
            "videos": os.path.join(os.path.expanduser("~"), "Videos")
        }

        self.websites = {
            "google": "https://www.google.com",
            "youtube": "https://www.youtube.com",
            "github": "https://www.github.com",
            "wikipedia": "https://www.wikipedia.org",
            "gmail": "https://www.gmail.com",
            "amazon": "https://www.amazon.com",
            "netflix": "https://www.netflix.com"
        }

    def open_website(self, url_or_name):
        try:
            if url_or_name.lower() in self.websites:
                url = self.websites[url_or_name.lower()]
                webbrowser.open(url)
                return True
            return False
        except Exception:
            return False

    def open_directory(self, dir_path):
        try:
            if dir_path.lower() in self.common_dirs:
                full_path = self.common_dirs[dir_path.lower()]
            else:
                full_path = os.path.expandvars(os.path.expanduser(dir_path))
            
            if os.path.exists(full_path):
                subprocess.Popen(f'explorer "{full_path}"')
                return True
            return False
        except Exception:
            return False

    def open_application(self, app_name):
        try:
            app_path = self.apps.get(app_name.lower())
            if app_path:
                subprocess.Popen(app_path)
                return True
            return False
        except Exception:
            return False