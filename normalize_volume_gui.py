import sys
import subprocess
import tkinter as tk
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

def install_packages():
    """Installs required packages if they are not already installed."""
    required = {'comtypes', 'pycaw'}
    installed = set()
    for package in required:
        try:
            __import__(package)
            installed.add(package)
        except ImportError:
            pass
    missing = required - installed
    if missing:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing])

install_packages()

def set_volume(level):
    """Sets the system's master volume to the specified level."""
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    
    # Normalize the level to be between 0.0 and 1.0
    normalized_level = max(0.0, min(level / 100.0, 1.0))
    volume.SetMasterVolumeLevelScalar(normalized_level, None)

def update_volume(value):
    """Updates the volume label and sets the volume when the slider changes."""
    desired_volume = int(value)
    label.config(text=f"Volume: {desired_volume}%")
    set_volume(desired_volume)

# Create the main window
root = tk.Tk()
root.title("Volume Normalizer")

# Set the size of the window
root.geometry("300x150")

# Create a label to display the current volume
label = tk.Label(root, text="Volume: 50%", font=('Helvetica', 16))
label.pack(pady=20)

# Create a slider to adjust the volume
volume_slider = tk.Scale(root, from_=0, to=100, orient='horizontal', command=update_volume)
volume_slider.set(50)  # Set default value to 50%
volume_slider.pack()

# Run the application
root.mainloop()
