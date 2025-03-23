import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from docx.oxml.ns import qn
from lxml.etree import Element as OxmlElement
from PIL import Image
import io

class Word2PPTX:
    def __init__(self, input_doc, output_ppt, template_ppt, max_leaf_count=8):
        self.input_doc = input_doc
        self.output_ppt = output_ppt
        self.template_ppt = template_ppt
        self.max_leaf_count = max_leaf_count
        self.doc = Document(input_doc)
        self.prs = Presentation(template_ppt)
        if not template_ppt:
            raise ValueError("A template PPT file must be provided")

    def count_leaf_headings(self, start_idx, end_idx):
        """统计指定范围内叶子标题数量（2-4级，不含子标题）"""
        leaf_count = 0
        i = start_idx
        while i < end_idx:
            para = self.doc.paragraphs[i]
            style = para.style.name
            if "Heading" in style:
                level = int(style.split()[-1]) if style.split()[-1].isdigit() else 0
                if 2 <= level <= 4:
                    is_leaf = True
                    for j in range(i + 1, end_idx):
                        next_style = self.doc.paragraphs[j].style.name
                        if "Heading" in next_style and int(next_style.split()[-1]) > level:
                            is_leaf = False
                            break
                    if is_leaf:
                        leaf_count += 1
            i += 1
        return leaf_count

    def extract_images(self, start_idx, end_idx):
        """提取指定范围内的图片，包括内联和浮动图片"""
        images = []
        print(f"Checking images between paragraphs {start_idx} and {end_idx}")
        
        inline_count = 0
        for idx, shape in enumerate(self.doc.inline_shapes):
            if shape.type == 3:
                inline_count += 1
                shape_element = shape._inline
                print(f"Found inline shape {idx} (type: picture)")
                for i in range(start_idx, end_idx):
                    para = self.doc.paragraphs[i]
                    print(f"  Checking paragraph {i}: '{para.text.strip()}'")
                    for run_idx, run in enumerate(para.runs):
                        if shape_element in run._element.getparent().getchildren():
                            image_rid = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
                            print(f"    Matched inline image in run {run_idx} of paragraph {i}")
                            try:
                                image_blob = self.doc.part.related_parts[image_rid].blob
                                images.append(image_blob)
                                print(f"    Successfully extracted inline image (RID: {image_rid})")
                                break
                            except KeyError:
                                print(f"    Warning: Inline image RID {image_rid} not found")
                                break
                    if images and images[-1] == image_blob:
                        break
        
        float_count = 0
        for rel in self.doc.part.rels.values():
            if "image" in rel.target_ref:
                float_count += 1
                print(f"Found potential floating image in relationships: {rel.target_ref}")
                for i in range(start_idx, end_idx):
                    para = self.doc.paragraphs[i]
                    for run in para.runs:
                        drawing_elements = run._element.findall(qn('w:drawing'))
                        for drawing in drawing_elements:
                            blip = drawing.find('.//' + qn('a:blip'))
                            if blip is not None and blip.embed == rel.rId:
                                print(f"  Matched floating image in paragraph {i}")
                                try:
                                    image_blob = self.doc.part.related_parts[rel.rId].blob
                                    images.append(image_blob)
                                    print(f"  Successfully extracted floating image (RID: {rel.rId})")
                                    break
                                except KeyError:
                                    print(f"  Warning: Floating image RID {rel.rId} not found")
                                    break
                        if images and images[-1] == image_blob:
                            break
        
        if not images:
            print(f"No images found in range {start_idx} to {end_idx}. Inline shapes: {inline_count}, Float relationships: {float_count}")
        return images

    def add_slide(self, layout_idx, title, subheadings=None, notes="", image_blob=None):
        """添加幻灯片，图片右下对齐，高或宽为PPT的一半，保持原始质量"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[layout_idx])
        
        # 设置标题
        slide.shapes.title.text = title
        
        # 添加正文内容
        content_shape = None
        for shape in slide.shapes:
            if shape.is_placeholder and shape.placeholder_format.idx == 1:
                content_shape = shape
                break
        
        if content_shape and subheadings:
            text_frame = content_shape.text_frame
            text_frame.clear()
            for i, subheading in enumerate(subheadings):
                p = text_frame.add_paragraph()
                p.text = subheading
                if i > 0:
                    p.level = 0
        
        # 处理图片：右下对齐，高或宽为PPT的一半，保持原始质量
        if image_blob:
            # 从字节数据加载图片，获取原始尺寸
            img = Image.open(io.BytesIO(image_blob))
            img_width, img_height = img.size
            img_aspect = img_width / img_height
            
            # 获取PPT尺寸（单位：EMU）
            ppt_width = self.prs.slide_width
            ppt_height = self.prs.slide_height
            max_width = ppt_width // 2  # PPT宽度的一半
            max_height = ppt_height // 2  # PPT高度的一半
            
            # 等比缩放显示尺寸
            if img_aspect > (max_width / max_height):  # 宽图，以宽度为基准
                display_width = max_width
                display_height = int(display_width / img_aspect)
            else:  # 高图，以高度为基准
                display_height = max_height
                display_width = int(display_height * img_aspect)
            
            # 计算右下对齐位置（单位：EMU）
            left = ppt_width - display_width
            top = ppt_height - display_height
            
            # 直接从字节数据插入图片，保持原始质量
            pic = slide.shapes.add_picture(io.BytesIO(image_blob), left, top, width=display_width, height=display_height)
            print(f"Inserted image into slide {len(self.prs.slides)-1}: size={img_width}x{img_height}, displayed as {display_width}x{display_height}, aligned bottom-right")
        
        if notes:
            slide.notes_slide.notes_text_frame.text = notes
        
        return slide

    def convert(self):
        """执行Word到PPT转换"""
        # 第一页：文章标题
        title = None
        title_idx = -1
        for i, para in enumerate(self.doc.paragraphs):
            if para.style.name == "Title" and para.text.strip():
                title = para.text.strip()
                title_idx = i
                break
        if not title and self.doc.paragraphs and self.doc.paragraphs[0].text.strip():
            title = self.doc.paragraphs[0].text.strip()
            title_idx = 0
        if title:
            if self.prs.slides:
                slide = self.prs.slides[0]
                slide.shapes.title.text = title
                print(f"Set title '{title}' to existing first slide")
            else:
                self.add_slide(0, title)
                print(f"Added title '{title}' to new first slide")
        else:
            if self.prs.slides:
                slide = self.prs.slides[0]
                slide.shapes.title.text = "Untitled Document"
                print("No Title found, set 'Untitled Document' to existing first slide")
            else:
                self.add_slide(0, "Untitled Document")
                print("No Title found, added 'Untitled Document' to new first slide")
            title_idx = -1

        # 收集Title和第一个Heading 1之间的文字，用于Agenda的备注
        agenda_notes = ""
        first_h1_idx = -1
        for i, para in enumerate(self.doc.paragraphs):
            if para.style.name == "Heading 1" and para.text.strip():
                first_h1_idx = i
                break
        if title_idx != -1 and first_h1_idx != -1 and first_h1_idx > title_idx + 1:
            for i in range(title_idx + 1, first_h1_idx):
                text = self.doc.paragraphs[i].text.strip()
                if text:
                    agenda_notes += text + "\n"
            if agenda_notes:
                print(f"Collected notes for Agenda between Title and first Heading 1:\n{agenda_notes.strip()}")

        # 第二页：Agenda（仅Heading 1）
        toc_subheadings = []
        for para in self.doc.paragraphs:
            if para.style.name == "Heading 1" and para.text.strip():
                toc_subheadings.append(para.text.strip())
        if toc_subheadings:
            if len(self.prs.slides) > 1:
                slide = self.prs.slides[1]
                slide.shapes.title.text = "Agenda"
                content_shape = None
                for shape in slide.shapes:
                    if shape.is_placeholder and shape.placeholder_format.idx == 1:
                        content_shape = shape
                        break
                if content_shape:
                    text_frame = content_shape.text_frame
                    text_frame.clear()
                    for i, subheading in enumerate(toc_subheadings):
                        p = text_frame.add_paragraph()
                        p.text = subheading
                        if i > 0:
                            p.level = 0
                if agenda_notes:
                    slide.notes_slide.notes_text_frame.text = agenda_notes.strip()
                print(f"Set Agenda with {len(toc_subheadings)} Heading 1 items to existing second slide")
            else:
                self.add_slide(1, "Agenda", toc_subheadings, agenda_notes)
                print(f"Added Agenda with {len(toc_subheadings)} Heading 1 items as new second slide")

        # 找到所有Heading 1的范围
        sections = []
        start_idx = 0
        for i, para in enumerate(self.doc.paragraphs):
            if "Heading 1" in para.style.name and i > 0:
                sections.append((start_idx, i))
                start_idx = i
        sections.append((start_idx, len(self.doc.paragraphs)))

        # 处理每个大章节
        slide_idx = 2
        for start_idx, end_idx in sections[1:]:
            title = self.doc.paragraphs[start_idx].text.strip()
            leaf_count = self.count_leaf_headings(start_idx, end_idx)
            images = self.extract_images(start_idx, end_idx)
            
            if leaf_count <= self.max_leaf_count and len(images) <= 1:
                notes = ""
                subheadings = []
                for i in range(start_idx + 1, end_idx):
                    para = self.doc.paragraphs[i]
                    style = para.style.name
                    text = para.text.strip()
                    if text:
                        notes += text + "\n"
                        if "Heading" in style and "Heading 1" not in style:
                            subheadings.append(text)
                if len(self.prs.slides) > slide_idx:
                    slide = self.prs.slides[slide_idx]
                    slide.shapes.title.text = title
                    content_shape = None
                    for shape in slide.shapes:
                        if shape.is_placeholder and shape.placeholder_format.idx == 1:
                            content_shape = shape
                            break
                    if content_shape and subheadings:
                        text_frame = content_shape.text_frame
                        text_frame.clear()
                        for i, subheading in enumerate(subheadings):
                            p = text_frame.add_paragraph()
                            p.text = subheading
                            if i > 0:
                                p.level = 0
                    if notes:
                        slide.notes_slide.notes_text_frame.text = notes
                    if images:
                        self.add_slide(1, title, subheadings, notes, images[0])  # 直接使用新逻辑
                else:
                    self.add_slide(1, title, subheadings, notes, images[0] if images else None)
                slide_idx += 1
            else:
                current_notes = ""
                current_subheadings = []
                current_images = []
                sub_start = start_idx + 1
                for i in range(start_idx + 1, end_idx):
                    para = self.doc.paragraphs[i]
                    style = para.style.name
                    text = para.text.strip()
                    if "Heading" in style and i > sub_start:
                        if current_notes or current_images:
                            if len(self.prs.slides) > slide_idx:
                                slide = self.prs.slides[slide_idx]
                                slide.shapes.title.text = title
                                content_shape = None
                                for shape in slide.shapes:
                                    if shape.is_placeholder and shape.placeholder_format.idx == 1:
                                        content_shape = shape
                                        break
                                if content_shape and current_subheadings:
                                    text_frame = content_shape.text_frame
                                    text_frame.clear()
                                    for i, subheading in enumerate(current_subheadings):
                                        p = text_frame.add_paragraph()
                                        p.text = subheading
                                        if i > 0:
                                            p.level = 0
                                if current_notes:
                                    slide.notes_slide.notes_text_frame.text = current_notes
                                if current_images:
                                    self.add_slide(1, title, current_subheadings, current_notes, current_images[0])
                            else:
                                self.add_slide(1, title, current_subheadings, current_notes, current_images[0] if current_images else None)
                            slide_idx += 1
                            current_notes = ""
                            current_subheadings = []
                            current_images = []
                        sub_start = i
                    if text:
                        current_notes += text + "\n"
                        if "Heading" in style and "Heading 1" not in style:
                            current_subheadings.append(text)
                    current_images.extend(self.extract_images(i, i + 1))
                    
                    if len(current_images) > 1:
                        if len(self.prs.slides) > slide_idx:
                            slide = self.prs.slides[slide_idx]
                            slide.shapes.title.text = title
                            content_shape = None
                            for shape in slide.shapes:
                                if shape.is_placeholder and shape.placeholder_format.idx == 1:
                                    content_shape = shape
                                    break
                            if content_shape and current_subheadings:
                                text_frame = content_shape.text_frame
                                text_frame.clear()
                                for i, subheading in enumerate(current_subheadings):
                                    p = text_frame.add_paragraph()
                                    p.text = subheading
                                    if i > 0:
                                        p.level = 0
                            if current_notes:
                                slide.notes_slide.notes_text_frame.text = current_notes
                            self.add_slide(1, title, current_subheadings, current_notes, current_images[0])
                        else:
                            self.add_slide(1, title, current_subheadings, current_notes, current_images[0])
                        slide_idx += 1
                        current_notes = ""
                        current_subheadings = []
                        current_images = current_images[1:]
                        sub_start = i + 1
                
                if current_notes or current_images:
                    if len(self.prs.slides) > slide_idx:
                        slide = self.prs.slides[slide_idx]
                        slide.shapes.title.text = title
                        content_shape = None
                        for shape in slide.shapes:
                            if shape.is_placeholder and shape.placeholder_format.idx == 1:
                                content_shape = shape
                                break
                        if content_shape and current_subheadings:
                            text_frame = content_shape.text_frame
                            text_frame.clear()
                            for i, subheading in enumerate(current_subheadings):
                                p = text_frame.add_paragraph()
                                p.text = subheading
                                if i > 0:
                                    p.level = 0
                        if current_notes:
                            slide.notes_slide.notes_text_frame.text = current_notes
                        if current_images:
                            self.add_slide(1, title, current_subheadings, current_notes, current_images[0])
                    else:
                        self.add_slide(1, title, current_subheadings, current_notes, current_images[0] if current_images else None)
                    slide_idx += 1

        self.prs.save(self.output_ppt)
        print(f"PPT saved as {self.output_ppt} with {len(self.prs.slides)} slides")
        return self.output_ppt