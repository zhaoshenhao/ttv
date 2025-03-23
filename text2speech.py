import os
from pptx import Presentation
from gtts import gTTS

class Text2Speech:
    def __init__(self, ppt_file, audio_dir, lang="zh-cn", speed="normal"):
        self.ppt_file = ppt_file
        self.audio_dir = audio_dir
        self.lang = lang
        self.speed = speed
        self.prs = Presentation(ppt_file)
        if not os.path.exists(audio_dir):
            os.makedirs(audio_dir)

    def convert(self):
        """将PPT Notes转为语音"""
        for i, slide in enumerate(self.prs.slides):
            audio_file = os.path.join(self.audio_dir, f"slide{i}.mp3")
            if os.path.exists(audio_file):
                print(f"Using pre-recorded audio for slide {i}: {audio_file}")
                continue
            
            text = slide.notes_slide.notes_text_frame.text
            if text:
                slow = True if self.speed == "slow" else False
                tts = gTTS(text, lang=self.lang, slow=slow)
                tts.save(audio_file)
                print(f"Generated audio for slide {i}: {audio_file}")