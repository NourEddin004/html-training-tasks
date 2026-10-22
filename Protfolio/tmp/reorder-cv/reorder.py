import sys
sys.path.insert(0,'/private/tmp/cv-pdf-tools')
import fitz
from pathlib import Path
src=fitz.open('/Users/noureddinsmacbook/Documents/My IT CV.pdf')
out=fitz.open()
W,H=595.28,841.89
# Each segment retains the source's actual text, font, colour and geometry.
def segment(dest, page, top, bottom, y):
 part=fitz.open(); part.insert_pdf(src,from_page=page,to_page=page)
 p=part[0]
 if top>0: p.add_redact_annot(fitz.Rect(0,0,W,top),fill=False)
 if bottom<H: p.add_redact_annot(fitz.Rect(0,bottom,W,H),fill=False)
 p.apply_redactions(images=0,graphics=0)
 clip=fitz.Rect(0,top,W,bottom)
 dest.show_pdf_page(fitz.Rect(0,y,W,y+bottom-top),part,0,clip=clip)
 for l in src[page].get_links():
  if clip.contains(l['from']) and l.get('uri'):
   r=l['from']; dest.insert_link({'kind':fitz.LINK_URI,'from':fitz.Rect(r.x0,r.y0-top+y,r.x1,r.y1-top+y),'uri':l['uri']})
 part.close()
 return y+bottom-top
p=out.new_page(width=W,height=H)
y=segment(p,0,40,235,40)
y=segment(p,2,298,370,y+20)
y=segment(p,2,385,512,y+20)
y=segment(p,1,192,349,y+20)
p=out.new_page(width=W,height=H)
y=segment(p,0,362,598,40)
y=segment(p,0,248,346,y+20)
y=segment(p,1,361,686,y+20)
p=out.new_page(width=W,height=H)
y=segment(p,1,698,773,40)
y=segment(p,2,20,227,y+16)
y=segment(p,2,240,285,y+20)
y=segment(p,0,610,764,y+20)
y=segment(p,1,21,177,y)
print('Last page bottom:',y)
out.set_metadata({'title':'NourEddin Abu Salhieh CV','author':'NourEddin Abu Salhieh'})
path='output/cv/My-IT-CV-Reordered.pdf'
out.save(path,garbage=4,deflate=True)
for i,p in enumerate(out): p.get_pixmap(matrix=fitz.Matrix(1.6,1.6)).save(f'tmp/reorder-cv/final-{i+1}.png')
