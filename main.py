from pptx import Presentation 
from pptx.util import Pt

# Opening file
prs = Presentation('assets/Both_Tracks_Golden_Certificate.pptx')
slide = prs.slides[0]
name = "Akhil Bhalerao"

for shape in slide.shapes:
    if shape.has_text_frame:
        text_frame = shape.text_frame
        
        if (text_frame.text == 'Name'):
            text_frame.clear()
            p = text_frame.paragraphs[0]
            run = p.add_run()
            run.text = name
            font = run.font
            font.name = 'Caveat'
            font.size = Pt(40)

prs.save(f"output/Output_{name}.pdf")
