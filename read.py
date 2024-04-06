from pptx import Presentation
from pyttsx3 import init
import time

def read_ppt(file_path):
    prs = Presentation(file_path)
    engine = init()

    for slide in prs.slides:
        shapes = sorted(slide.shapes, key=lambda shape: (shape.top, shape.left))
        for shape in shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text = paragraph.text
                    if text.strip():
                        engine.say(text)
                        engine.runAndWait()
                        time.sleep(0.5)

if __name__ == "__main__":
    ppt_file = "sample1.pptx"
    read_ppt(ppt_file)