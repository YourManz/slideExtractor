import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess

def resource_path(relative_path):
    """ Get absolute path to resource (used for PyInstaller binary access) """
    try:
        base_path = sys._MEIPASS  # PyInstaller uses this
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def select_video():
    filepath = filedialog.askopenfilename(
        title="Select Video",
        filetypes=[("Video files", "*.mp4 *.mov *.avi *.mkv")]
    )
    if filepath:
        video_path.set(filepath)
        status.set("Ready")

def extract_slides():
    path = video_path.get()
    if not path:
        status.set("Please select a video file.")
        return

    try:
        threshold = float(threshold_val.get())
    except ValueError:
        status.set("Invalid threshold.")
        return

    basename = os.path.splitext(os.path.basename(path))[0]
    out_dir = f"{basename}_slides"
    os.makedirs(out_dir, exist_ok=True)

    ffmpeg_path = resource_path("ffmpeg.exe")
    output_pattern = os.path.join(out_dir, "%04d.jpg")
    cmd = [
        ffmpeg_path, "-i", path,
        "-filter_complex", f"select=gt(scene\\,{threshold})",
        "-vsync", "vfr", output_pattern
    ]

    try:
        status.set("Extracting...")
        root.update()
        subprocess.run(cmd, check=True)
        status.set(f"Done! Saved to: {out_dir}")
        os.startfile(out_dir)
    except Exception as e:
        status.set(f"Error: {e}")

# GUI Setup
root = tk.Tk()
root.title("Slide Extractor")
root.geometry("400x200")
video_path = tk.StringVar()
threshold_val = tk.StringVar(value="0.2")
status = tk.StringVar(value="Select a video.")

tk.Label(root, text="Video Path:").pack()
tk.Entry(root, textvariable=video_path, width=50).pack()
tk.Button(root, text="Browse", command=select_video).pack(pady=5)

tk.Label(root, text="Scene Threshold (e.g. 0.2):").pack()
tk.Entry(root, textvariable=threshold_val, width=10).pack()

tk.Button(root, text="Extract Slides", command=extract_slides).pack(pady=10)
tk.Label(root, textvariable=status, fg="blue").pack()

root.mainloop()