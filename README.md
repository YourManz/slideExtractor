# Slide Extractor

A small GUI tool to grab slides or frames from a video using `ffmpeg` and
export them to PDF or PowerPoint.

## Features

- Scene change detection with configurable threshold.
- Optional capture at specific timestamps.
- Custom output directory selection.
- Thumbnail preview of the first slide.
- Progress bar while extraction runs.
- Export to **PPTX** or **PDF** with an option to remove the intermediate JPGs.

## Usage

Install the required dependencies:

```bash
pip install -r requirements.txt
```

Run the script directly with Python:

```bash
python slide_extractor.py
```

To package a standalone executable with PyInstaller:

```bash
python -m PyInstaller --onefile --noconsole --add-binary "ffmpeg.exe;." slide_extractor.py
```
