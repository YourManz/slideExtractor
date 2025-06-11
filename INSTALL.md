# Installation

Follow these steps to set up **Slide Extractor** from source.

## 1. Clone the repository

```bash
git clone <repo-url>
cd slideExtractor
```

Alternatively, download a source release and unpack it.

## 2. Install Python dependencies

Ensure Python 3 is installed on your system, then install the required
packages:

```bash
pip install -r requirements.txt
```

## 3. Install FFmpeg

On Windows, the repository includes a prebuilt `ffmpeg.exe`. For other
platforms, install FFmpeg separately and make sure it is available on your
`PATH` as `ffmpeg`.

## 4. Run the application

Start the GUI directly with Python:

```bash
python slide_extractor.py
```

To build a standalone executable using PyInstaller:

```bash
python -m PyInstaller --onefile --noconsole --add-binary "ffmpeg.exe;." slide_extractor.py
```

The resulting executable will appear in the `dist` directory.
