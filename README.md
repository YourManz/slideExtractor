# Slide Extractor

## Packaging

Install the required dependencies first:

```bash
pip install -r requirements.txt
```

Then build the executable:

```bash
python -m PyInstaller --onefile --noconsole --add-binary "ffmpeg.exe;." slide_extractor.py
```
