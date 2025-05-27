# prez_maker: Automated Presentation Generator with Google Images

## Overview

`makePrezWithGoogledImages.py` is a Python script that automatically generates a PowerPoint presentation (`.pptx`) from structured slide data and enriches each slide with relevant images fetched from Google Images. The slide content (titles, text, keywords, notes) is defined externally in a `slides.json` file, making it easy to update or reuse content for new presentations.

## Features
- **Automated Slide Creation**: Generates slides based on structured JSON data.
- **Image Integration**: Downloads and inserts relevant images from Google Images for each slide using the specified keywords.
- **Presenter Notes**: Adds presenter notes to each slide.
- **Easy Content Management**: All slide content is maintained in `slides.json`.
- **Cleanup**: Removes downloaded images after the presentation is generated.

## Requirements
- Python 3.7+
- [python-pptx](https://python-pptx.readthedocs.io/)
- [icrawler](https://github.com/hellock/icrawler)
- [Pillow (PIL)](https://pillow.readthedocs.io/en/stable/)

You can install the dependencies with:
```bash
pip install python-pptx icrawler Pillow
```

## Usage
1. **Prepare Slide Data**: Edit `slides.json` to define your slides. Each slide should include a `title`, `text`, `keywords` (for image search), and `notes`.
2. **Run the Script**:
   ```bash
   python3 makePrezWithGoogledImages.py
   ```
3. **Output**: The script will generate a PowerPoint file named `presentation_<timestamp>.pptx` in the current directory.

## How It Works
- The script reads `slides.json` for slide content.
- For each slide, it uses the `keywords` field to search Google Images and downloads up to 5 relevant images.
- The first suitable image is inserted into the slide.
- Presenter notes are added from the `notes` field.
- Temporary images are deleted after the presentation is saved.

## Customization
- **Change Slide Content**: Edit `slides.json` as needed.
- **Image Search Tuning**: Adjust the `keywords` field in each slide for better image results.
- **Output Filename**: The script auto-generates a filename with a timestamp, but you can modify this in the script if desired.

## Limitations
- Google Images scraping is subject to rate limits and may occasionally fail to download some images.
- The script does not support advanced slide layouts or animations.

## License
MIT License. See [LICENSE](LICENSE) for details.

## Author
Created by [Your Name].

---

**Note:** This README describes only `makePrezWithGoogledImages.py`. Other scripts in this repository are not covered here.
