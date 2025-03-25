# Word文章到视频转换器

作为一个业余 YouTube 科普视频发布者和兼职 IT 培训师，我经常需要将自己撰写的博客和教程制作成视频。这是一项非常耗时的任务，因此我开发了这个小工具。

## 这是一个 Word 文章到视频的转换工具

它的流程如下：

Word -> PPT -> Audio -> Video。

文字到音频部分，我使用了开源的 **F5-TTS**，感谢 **F5-TTS** 的作者。  
音频到视频部分，我使用 **pywin32** 和 **PowerPoint** 保存每个幻灯片页面，然后利用 **moviepy** 制作视频和字幕。

## 命令行

该工具集成在一个命令行中，**必须在 Windows 平台上运行**。

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

## 安装

1. **建议使用 Conda 创建 Python 虚拟环境**。由于需要 F5-TTS 支持，请按照它的安装指南进行安装。
2. **完成 F5-TTS 安装后，安装本项目的依赖包**。

```bash
conda create -n f5-tts python=3.10
conda activate f5-tts
# 安装适用于您的 CUDA 版本的 PyTorch，例如：
pip install torch==2.4.0+cu124 torchaudio==2.4.0+cu124 --extra-index-url https://download.pytorch.org/whl/cu124
pip install f5-tts
git clone https://github.com/zhaoshenhao/ttv.git
cd ttv
pip install -r requirements.txt
```

## Word 格式约定

1. 必须使用 Office 默认格式。
2. 必须包含 **Title**（标题）。
3. 必须使用 **Headings**（从 1 到 4 级）。
4. 可以插入图片。

## 工具制作视频和字幕的规则

1. **工具使用 Title 生成第一页**，并生成**无声视频 2 秒**。
2. **工具收集 Heading 1 作为 Agenda**，并使用 **Title 和第一个 Heading 1 之间的文本**来解说 Agenda。
3. **每个 Heading 1 代表一个大章节（Section）**，其子 Heading（2-4 级）将被用于该章节的 Agenda 生成。
4. **如果章节中没有图片**，该章节的内容将用于解说该章节的 Agenda。
5. **如果章节中有图片**：
   - 每张图片都会单独生成一个幻灯片。
   - 图片前面的**Heading 和图片之间的文本**用于解说该图片所在章节的 Agenda。
   - **图片后的文本**用于解说图片内容。
6. **如果图片后面仍有内容**：
   - 章节的 Agenda 会被重新添加，并使用剩余内容进行解说。
   - **重复步骤 4-6 直到该章节结束**。
7. **所有文字会按中英文的逗号、句号、分号、冒号分割成句子**，字幕按照句子组织。