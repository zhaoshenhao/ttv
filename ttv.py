import argparse
from word2pptx import Word2PPTX
from text2speech import Text2Speech
from ppt2video import PPT2Video

def main():
    parser = argparse.ArgumentParser(
        description="Convert Word to Video with customizable options",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to execute", required=True)

    # word2ppt 命令
    parser_word2ppt = subparsers.add_parser("word2ppt", help="Convert Word to PPT")
    parser_word2ppt.add_argument("-w", "--word", default="input.docx", help="Input Word document")
    parser_word2ppt.add_argument("-p", "--ppt", default="output.pptx", help="Output PPT file")
    parser_word2ppt.add_argument("-t", "--template", required=True, help="PPT template file")
    parser_word2ppt.add_argument("-m", "--max-leaf-count", type=int, default=8, help="Max leaf headings before splitting")

    # tts 命令
    parser_tts = subparsers.add_parser("tts", help="Convert PPT notes to speech")
    parser_tts.add_argument("-p", "--ppt", default="output.pptx", help="Input PPT file")
    parser_tts.add_argument("-o", "--output-dir", default="audio", help="Directory for output files")
    parser_tts.add_argument("-l", "--lang", default="zh", choices=["zh", "en"], help="Language (zh or en)")

    # ppt2video 命令
    parser_ppt2video = subparsers.add_parser("ppt2video", help="Convert PPT to video")
    parser_ppt2video.add_argument("-p", "--ppt", default="output.pptx", help="Input PPT file")
    parser_ppt2video.add_argument("-v", "--video", default="output.mp4", help="Output video file")
    parser_ppt2video.add_argument("-o", "--output-dir", default="audio", help="Directory for output files")

    # all 命令
    parser_all = subparsers.add_parser("all", help="Run all steps: Word to PPT, TTS, and PPT to video")
    parser_all.add_argument("-w", "--word", default="input.docx", help="Input Word document")
    parser_all.add_argument("-p", "--ppt", default="output.pptx", help="Output PPT file")
    parser_all.add_argument("-v", "--video", default="output.mp4", help="Output video file")
    parser_all.add_argument("-o", "--output-dir", default="audio", help="Directory for output file")
    parser_all.add_argument("-l", "--lang", default="zh", choices=["zh", "en"], help="Language (zh or en)")
    parser_all.add_argument("-t", "--template", required=True, help="PPT template file")
    parser_all.add_argument("-m", "--max-leaf-count", type=int, default=8, help="Max leaf headings before splitting")

    # 解析参数
    args = parser.parse_args()

    # 根据命令执行相应逻辑
    if args.command == "word2ppt":
        converter = Word2PPTX(args.word, args.ppt, args.template, args.max_leaf_count)
        converter.convert()
    elif args.command == "tts":
        converter = Text2Speech(args.ppt, args.lang)
        converter.convert()
    elif args.command == "ppt2video":
        converter = PPT2Video(args.ppt, args.video)
        converter.convert()
    elif args.command == "all":
        word2ppt = Word2PPTX(args.word, args.ppt, args.template, args.max_leaf_count)
        word2ppt.convert()
        tts = Text2Speech(args.ppt, args.lang)
        tts.convert()
        ppt2video = PPT2Video(args.ppt, args.video)
        ppt2video.convert()

if __name__ == "__main__":
    main()