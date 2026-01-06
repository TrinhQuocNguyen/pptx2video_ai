# PPTX to Video Using AI

This repo converts a `.pptx` file to a video, and read out speaker's notes along slides using text to speech technique.

## Usage

### Installation 

1) Install required packages
```
pip install -r requirements.txt
```

You may need **"poppler-windows"** and **"ffmpeg"**:
```
# Download from https://github.com/oschwartz10612/poppler-windows/releases/
POPPLER_PATH = r'C:\Users\nguye\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin' # Update this path

# Download ffmpeg from https://github.com/BtbN/FFmpeg-Builds/releases
FFMPEG_PATH = r"C:\ffmpeg\bin\ffmpeg.exe"
```

2) Convert/save the PPTX file into PDF file
   (Export option)

3) Run the command to generate the video:
```
python ppt_presenter.py --pptx example/swot.pptx --pdf example/swot.pdf -o example/swot.mp4
```

### The pipeline

- PPTX -> PDF => IMAGE
- PPTX Speaker's Note => AUDIO
=> **VIDEO = IMAGE + AUDIO**
