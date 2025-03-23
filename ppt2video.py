import os
from pptx import Presentation
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips
import win32com.client  # 需要 pywin32: pip install pywin32
import shutil

class PPT2Video:
    def __init__(self, ppt_file, video_file, audio_dir):
        self.ppt_file = ppt_file
        self.video_file = video_file
        self.audio_dir = audio_dir
        self.prs = Presentation(ppt_file)
        self.default_duration = 5  # 默认无声视频时长（秒）
        self.temp_dir = "temp_slides"

    def export_slides_to_images(self):
        """使用PowerPoint将每页幻灯片导出为图片，确保顺序正确"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)  # 清理旧的临时文件夹
        os.makedirs(self.temp_dir)

        # 启动PowerPoint应用
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(os.path.abspath(self.ppt_file))
        
        # 导出所有幻灯片为PNG
        presentation.Export(os.path.abspath(self.temp_dir), "PNG")
        presentation.Close()
        powerpoint.Quit()
        
        # 获取导出的文件并按幻灯片顺序重命名
        exported_files = sorted(
            [f for f in os.listdir(self.temp_dir) if f.lower().endswith(".png")],
            key=lambda x: int(x.split('Slide')[1].split('.')[0])  # 按Slide编号排序
        )
        slide_images = []
        for i, old_file in enumerate(exported_files):
            old_path = os.path.join(self.temp_dir, old_file)
            new_path = os.path.join(self.temp_dir, f"slide_{i}.png")
            os.rename(old_path, new_path)
            print(f"Generated image for slide {i}: {new_path}")
            slide_images.append(new_path)
        
        return slide_images

    def convert(self):
        """将PPT和音频合成为视频"""
        # 导出幻灯片为图片
        slide_images = self.export_slides_to_images()
        
        if len(slide_images) != len(self.prs.slides):
            print(f"Warning: Number of exported images ({len(slide_images)}) does not match slides ({len(self.prs.slides)})")
        
        clips = []
        for i, slide in enumerate(self.prs.slides):
            if i >= len(slide_images):
                print(f"Error: No image for slide {i}, stopping")
                break
            img_file = slide_images[i]
            audio_file = os.path.join(self.audio_dir, f"slide{i}.mp3")
            
            # 创建剪辑
            if os.path.exists(audio_file):
                audio = AudioFileClip(audio_file)
                clip = ImageClip(img_file, duration=audio.duration)
                try:
                    clip = clip.set_audio(audio)  # 新版API
                except AttributeError:
                    clip.audio = audio  # 旧版兼容
                print(f"Created clip for slide {i} with audio, duration {audio.duration} seconds")
            else:
                clip = ImageClip(img_file, duration=self.default_duration)
                print(f"Created silent clip for slide {i}, default duration {self.default_duration} seconds")
            
            clips.append(clip)
        
        if clips:
            final_video = concatenate_videoclips(clips)
            final_video.write_videofile(self.video_file, fps=24)
            print(f"Video saved as {self.video_file}")
            # 清理临时文件夹
            shutil.rmtree(self.temp_dir)
        else:
            print("No clips to process, video not generated")