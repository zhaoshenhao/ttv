import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from docx.oxml.ns import qn
from lxml.etree import Element as OxmlElement

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
        """添加幻灯片，使用默认文字框，图片置于底层"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[layout_idx])
        
        # 先添加图片（若有），确保在底层
        if image_blob:
            image_path = f"slide_image_{len(self.prs.slides)-1}.png"
            with open(image_path, "wb") as f:
                f.write(image_blob)
            left = top = Inches(0)
            width = self.prs.slide_width
            height = self.prs.slide_height
            pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)
            # 将图片移到最底层
            slide.shapes._spTree.remove(pic._element)
            slide.shapes._spTree.insert(2, pic._element)  # 插入到靠近底部的索引（2是标题和正文之后）
            print(f"Added image as bottom-layer shape for slide {len(self.prs.slides)-1}: {image_path}")
        
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
        
        if notes:
            slide.notes_slide.notes_text_frame.text = notes
        
        return slide

    def convert(self):
        """执行Word到PPT转换"""
        # 第一页：文章标题
        title = None
        for para in self.doc.paragraphs:
            if para.style.name == "Title" and para.text.strip():
                title = para.text.strip()
                break
        if not title and self.doc.paragraphs and self.doc.paragraphs[0].text.strip():
            title = self.doc.paragraphs[0].text.strip()
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

        # 第二页：Agenda（仅Heading 1）
        toc_subheadings = []
        for para in self.doc.paragraphs:
            if para.style.name == "Heading 1" and para.text.strip():  # 修改为只收集Heading 1
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
                print(f"Set Agenda with {len(toc_subheadings)} Heading 1 items to existing second slide")
            else:
                self.add_slide(1, "Agenda", toc_subheadings)
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
                        image_path = f"slide_image_{slide_idx}.png"
                        with open(image_path, "wb") as f:
                            f.write(images[0])
                        left = top = Inches(0)
                        width = self.prs.slide_width
                        height = self.prs.slide_height
                        pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)
                        slide.shapes._spTree.remove(pic._element)
                        slide.shapes._spTree.insert(2, pic._element)
                        print(f"Added image as bottom-layer shape for slide {slide_idx}: {image_path}")
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
                                    image_path = f"slide_image_{slide_idx}.png"
                                    with open(image_path, "wb") as f:
                                        f.write(current_images[0])
                                    left = top = Inches(0)
                                    width = self.prs.slide_width
                                    height = self.prs.slide_height
                                    pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)
                                    slide.shapes._spTree.remove(pic._element)
                                    slide.shapes._spTree.insert(2, pic._element)
                                    print(f"Added image as bottom-layer shape for slide {slide_idx}: {image_path}")
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
                            image_path = f"slide_image_{slide_idx}.png"
                            with open(image_path, "wb") as f:
                                f.write(current_images[0])
                            left = top = Inches(0)
                            width = self.prs.slide_width
                            height = self.prs.slide_height
                            pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)
                            slide.shapes._spTree.remove(pic._element)
                            slide.shapes._spTree.insert(2, pic._element)
                            print(f"Added image as bottom-layer shape for slide {slide_idx}: {image_path}")
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
                            image_path = f"slide_image_{slide_idx}.png"
                            with open(image_path, "wb") as f:
                                f.write(current_images[0])
                            left = top = Inches(0)
                            width = self.prs.slide_width
                            height = self.prs.slide_height
                            pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)
                            slide.shapes._spTree.remove(pic._element)
                            slide.shapes._spTree.insert(2, pic._element)
                            print(f"Added image as bottom-layer shape for slide {slide_idx}: {image_path}")
                    else:
                        self.add_slide(1, title, current_subheadings, current_notes, current_images[0] if current_images else None)
                    slide_idx += 1

        self.prs.save(self.output_ppt)
        print(f"PPT saved as {self.output_ppt} with {len(self.prs.slides)} slides")
        return self.output_ppt