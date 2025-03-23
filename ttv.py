import argparse
from word2pptx import Word2PPTX
from text2speech import Text2Speech
from ppt2video import PPT2Video

def main():
    # 创建主解析器
    parser = argparse.ArgumentParser(
        description="Convert Word to Video with customizable options",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to execute", required=True)

    # word2ppt 命令
    parser_word2ppt = subparsers.add_parser("word2ppt", help="Convert Word to PPT")
    parser_word2ppt.add_argument("--input", default="input.docx", help="Input Word document")
    parser_word2ppt.add_argument("--ppt", default="output.pptx", help="Output PPT file")
    parser_word2ppt.add_argument("--template", required=True, help="PPT template file")
    parser_word2ppt.add_argument("--max-leaf-count", type=int, default=8, help="Max leaf headings before splitting")

    # tts 命令
    parser_tts = subparsers.add_parser("tts", help="Convert PPT notes to speech")
    parser_tts.add_argument("--ppt", default="output.pptx", help="Input PPT file")
    parser_tts.add_argument("--audio-dir", default="audio", help="Directory for pre-recorded audio")
    parser_tts.add_argument("--lang", default="zh-cn", help="TTS language")
    parser_tts.add_argument("--speed", default="normal", choices=["slow", "normal", "fast"], help="TTS speed")

    # ppt2video 命令
    parser_ppt2video = subparsers.add_parser("ppt2video", help="Convert PPT to video")
    parser_ppt2video.add_argument("--ppt", default="output.pptx", help="Input PPT file")
    parser_ppt2video.add_argument("--video", default="output.mp4", help="Output video file")
    parser_ppt2video.add_argument("--audio-dir", default="audio", help="Directory for audio files")

    # all 命令
    parser_all = subparsers.add_parser("all", help="Run all steps: Word to PPT, TTS, and PPT to video")
    parser_all.add_argument("--input", default="input.docx", help="Input Word document")
    parser_all.add_argument("--ppt", default="output.pptx", help="Output PPT file")
    parser_all.add_argument("--video", default="output.mp4", help="Output video file")
    parser_all.add_argument("--audio-dir", default="audio", help="Directory for pre-recorded audio")
    parser_all.add_argument("--lang", default="zh-cn", help="TTS language")
    parser_all.add_argument("--speed", default="normal", choices=["slow", "normal", "fast"], help="TTS speed")
    parser_all.add_argument("--template", required=True, help="PPT template file")
    parser_all.add_argument("--max-leaf-count", type=int, default=8, help="Max leaf headings before splitting")

    # 解析参数
    args = parser.parse_args()

    # 根据命令执行相应逻辑
    if args.command == "word2ppt":
        converter = Word2PPTX(args.input, args.ppt, args.template, args.max_leaf_count)
        converter.convert()
    elif args.command == "tts":
        converter = Text2Speech(args.ppt, args.audio_dir, args.lang, args.speed)
        converter.convert()
    elif args.command == "ppt2video":
        converter = PPT2Video(args.ppt, args.video, args.audio_dir)
        converter.convert()
    elif args.command == "all":
        word2ppt = Word2PPTX(args.input, args.ppt, args.template, args.max_leaf_count)
        word2ppt.convert()
        tts = Text2Speech(args.ppt, args.audio_dir, args.lang, args.speed)
        tts.convert()
        ppt2video = PPT2Video(args.ppt, args.video, args.audio_dir)
        ppt2video.convert()

if __name__ == "__main__":
    main()