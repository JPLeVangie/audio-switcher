import pystray
from PIL import Image
import ctypes
import os
import sys
import win32com.client
from pycaw.pycaw import AudioUtilities

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        print(f"Error checking admin status: {e}")
        return False

class AudioSwitcher:
    def __init__(self):
        print("Initializing AudioSwitcher...")
        self.devices = self.get_audio_devices()
        self.current_device_index = 0
        self.speaker_icon = self.load_icon('speaker_icon.png')
        self.headphone_icon = self.load_icon('headphone_icon.png')
        self.icon = self.create_icon()

    def get_audio_devices(self):
        try:
            print("Enumerating audio devices...")
            return AudioUtilities.GetAllDevices()
        except Exception as e:
            print(f"Error getting audio devices: {e}")
            return []

    def load_icon(self, filename):
        try:
            script_dir = os.path.dirname(os.path.realpath(__file__))
            icon_path = os.path.join(script_dir, filename)
            print(f"Loading icon: {icon_path}")
            return Image.open(icon_path)
        except Exception as e:
            print(f"Error loading icon '{filename}': {e}")
            return None

    def switch_audio_device(self):
        if not self.devices:
            print("No playback devices found.")
            return

        try:
            self.current_device_index = (self.current_device_index + 1) % len(self.devices)
            device = self.devices[self.current_device_index]
            self.set_default_audio_device(device.id)
            self.update_icon(device.FriendlyName)
            print(f"Switched to {device.FriendlyName}")
        except Exception as e:
            print(f"Error switching audio device: {e}")

    def set_default_audio_device(self, device_id):
        try:
            print(f"Setting default audio device to {device_id}...")
            winmm = win32com.client.Dispatch("WScript.Shell")
            winmm.SendKeys(chr(0xAF))  # Simulate key press to ensure audio change is recognized
            
            # Set default audio device
            ps_command = f'$devices = Get-AudioDevice -Playback;'
            ps_command += f'$deviceToSet = $devices | Where-Object {{ $_.ID -eq "{device_id}" }};'
            ps_command += f'if ($deviceToSet) {{ $deviceToSet | Set-AudioDevice -Verbose }} else {{ Write-Host "Device not found" }}'
            
            os.system(f'powershell -Command "{ps_command}"')
        except Exception as e:
            print(f"Error setting default audio device: {e}")

    def update_icon(self, device_name):
        try:
            print(f"Updating icon for device: {device_name}")
            if "speaker" in device_name.lower():
                self.icon.icon = self.speaker_icon
            elif "headphone" in device_name.lower():
                self.icon.icon = self.headphone_icon
            self.icon.update_menu()
        except Exception as e:
            print(f"Error updating icon: {e}")

    def create_icon(self):
        try:
            print("Creating system tray icon...")
            return pystray.Icon("audio_switcher", self.speaker_icon, "Audio Device Switcher", menu=pystray.Menu(
                pystray.MenuItem("Switch Audio Device", self.switch_audio_device),
                pystray.MenuItem("Exit", self.on_exit)
            ))
        except Exception as e:
            print(f"Error creating system tray icon: {e}")
            return None

    def setup(self, icon):
        icon.visible = True

    def on_click(self, icon, event):
        if event.button == pystray.MouseButton.LEFT:
            self.switch_audio_device()

    def on_exit(self, icon, item):
        icon.stop()

    def run(self):
        if self.icon:
            print("Running AudioSwitcher...")
            self.icon.run(setup=self.setup)
        else:
            print("Failed to create system tray icon, exiting.")

if __name__ == "__main__":
    if is_admin():
        try:
            audio_switcher = AudioSwitcher()
            audio_switcher.icon.on_click = audio_switcher.on_click
            audio_switcher.run()
        except Exception as e:
            print(f"An error occurred: {e}")
            input("Press Enter to close...")  # Keep window open for debugging
    else:
        print("Requesting administrator privileges...")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)