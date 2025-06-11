import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import glob

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

    global out_dir
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

# Export functions
def export_to_pptx(directory):
    if not directory or not os.path.isdir(directory):
        status.set("No slides to export.")
        return
    images = sorted(glob.glob(os.path.join(directory, "*.jpg")))
    if not images:
        status.set("No slides to export.")
        return
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    for img in images:
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            img,
            Inches(0),
            Inches(0),
            width=prs.slide_width,
            height=prs.slide_height,
        )
    basename = os.path.basename(directory)
    pptx_path = os.path.join(directory, f"{basename}.pptx")
    try:
        prs.save(pptx_path)
        status.set(f"Exported to {pptx_path}")
        os.startfile(pptx_path)
    except Exception as e:
        status.set(f"Error: {e}")


def export_to_pdf(directory):
    if not directory or not os.path.isdir(directory):
        status.set("No slides to export.")
        return
    images = sorted(glob.glob(os.path.join(directory, "*.jpg")))
    if not images:
        status.set("No slides to export.")
        return
    img_objs = [Image.open(p).convert("RGB") for p in images]
    basename = os.path.basename(directory)
    pdf_path = os.path.join(directory, f"{basename}.pdf")
    try:
        img_objs[0].save(pdf_path, save_all=True, append_images=img_objs[1:])
        status.set(f"Exported to {pdf_path}")
        os.startfile(pdf_path)
    except Exception as e:
        status.set(f"Error: {e}")
    finally:
        for img in img_objs:
            img.close()

# GUI Setup
root = tk.Tk()
root.title("Slide Extractor")
root.geometry("400x200")
video_path = tk.StringVar()
threshold_val = tk.StringVar(value="0.2")
status = tk.StringVar(value="Select a video.")
out_dir = ""

tk.Label(root, text="Video Path:").pack()
tk.Entry(root, textvariable=video_path, width=50).pack()
tk.Button(root, text="Browse", command=select_video).pack(pady=5)

tk.Label(root, text="Scene Threshold (e.g. 0.2):").pack()
tk.Entry(root, textvariable=threshold_val, width=10).pack()

tk.Button(root, text="Extract Slides", command=extract_slides).pack(pady=10)
tk.Button(root, text="Export to PPTX", command=lambda: export_to_pptx(out_dir)).pack()
tk.Button(root, text="Export to PDF", command=lambda: export_to_pdf(out_dir)).pack()
tk.Label(root, textvariable=status, fg="blue").pack()

root.mainloop()