# Word to Video Converter

As an amateur YouTube science content creator and a part-time IT trainer, I often need to turn my blogs and tutorials into videos. This is a time-consuming task, so I developed this tool.

## A Word-to-Video Conversion Tool

The process follows these steps:

Word -> PPT -> Audio -> Video.

For text-to-speech, I use the open-source **F5-TTS**—special thanks to the authors of **F5-TTS**.  
For audio-to-video, I use **pywin32** and **PowerPoint** to save each slide as an image, then **moviepy** to create the video and subtitles.

## Command Line Usage

This tool runs as a command-line utility and **must be used on a Windows platform**.

```bash
usage: ttv.py [-h] {word2ppt,tts,ppt2video,all} ...

Convert Word to Video with customizable options

positional arguments:
  {word2ppt,tts,ppt2video,all}
                        Command to execute
    word2ppt            Convert Word to PPT
    tts                 Convert PPT notes to speech
    ppt2video           Convert PPT to video
    all                 Run all steps: Word to PPT, TTS, and PPT to video

options:
  -h, --help            show this help message and exit
```

## Installation

1. **It is recommended to use Conda to create a Python virtual environment**. Since F5-TTS is required, please follow its installation guide.
2. **After installing F5-TTS, install the dependencies for this project**.

```bash
conda create -n f5-tts python=3.10
conda activate f5-tts
# Install PyTorch with your CUDA version, e.g.:
pip install torch==2.4.0+cu124 torchaudio==2.4.0+cu124 --extra-index-url https://download.pytorch.org/whl/cu124
pip install f5-tts
git clone https://github.com/zhaoshenhao/ttv.git
cd ttv
pip install -r requirements.txt
```

## Word Formatting Requirements

1. The document must use **Microsoft Office default formatting**.
2. It must contain a **Title**.
3. It must use **Headings**, from level 1 to level 4.
4. Images can be inserted.

## Video and Subtitle Generation Rules

1. **The tool generates the first slide from the Title** and creates a **2-second silent video**.
2. **Heading 1 is collected as the Agenda**, and the text between **Title and the first Heading 1** is used to explain the Agenda.
3. **Each Heading 1 represents a major section**, and its subheadings (Heading 2-4) are used to generate the section’s Agenda.
4. **If a section has no images**, the text in the section will be used to explain the section’s Agenda.
5. **If a section contains images**:
   - Each image generates a separate slide.
   - The **text before the image and after the nearest preceding Heading** is used to explain the section’s Agenda.
   - The **text after the image** is used to describe the image.
6. **If there is additional text after an image**:
   - The section’s Agenda is repeated, and the remaining content is used for explanation.
   - **Steps 4-6 are repeated until the section is complete**.
7. **All text is split into sentences based on commas, periods, semicolons, and colons in both English and Chinese**. Subtitles are organized by sentence.
