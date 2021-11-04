from pptx import Presentation 
from pptx.util import Pt
import os
import comtypes.client

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

prs.save(f"output/Output_{name}.pptx")

input_file_path = os.path.abspath(f"output/Output_{name}.pptx")
output_file_path = os.path.abspath(f"output/Output_{name}.pdf")

#%% Create powerpoint application object
powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

#%% Set visibility to minimize
powerpoint.Visible = 1

#%% Open the powerpoint slides
slides = powerpoint.Presentations.Open(input_file_path)

#%% Save as PDF (formatType = 32)
slides.SaveAs(output_file_path, 32)

#%% Close the slide deck
slides.Close()