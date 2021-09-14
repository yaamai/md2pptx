from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE


pres = Presentation("base.pptx")
print(pres.core_properties.__dict__)
print(pres.slide_master)
dir(pres.slide_master)
print(pres.slide_master.slide_layouts)
print(len(pres.slide_master.slide_layouts))
print(pres.slide_master.slide_layouts[0].shapes)
print(len(pres.slide_master.slide_layouts[0].shapes))

for layout in pres.slide_master.slide_layouts:
    print(len(layout.shapes))
    for shape in layout.shapes:
        print(shape)
        if hasattr(shape, "type"):
            print(shape.type)
        print(shape.left, shape.top, shape.width, shape.height)


print("---")
print(pres.slide_width, pres.slide_height)

print("---")
for shape in pres.slide_master.shapes:
    print(shape)
    if hasattr(shape, "type"):
        print(shape.type)
    print(shape.left, shape.top, shape.width, shape.height)

print("---")
for placeholder in pres.slide_master.placeholders:
    print(placeholder)
    print(placeholder.left, placeholder.top, placeholder.width, placeholder.height)


import drawSvg as draw

print("---")
print(pres.slide_width, pres.slide_height)

# drawSvg was designed for the coordinates x and y to increase rightward and upward. (#11)
d = draw.Drawing(pres.slide_width, pres.slide_height)
g = draw.Group(transform="scale(1,-1) translate(0,{})".format(pres.slide_height))

for shape in pres.slide_master.shapes:
    print(shape.shape_type, shape.left, shape.top, shape.width, shape.height)
    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        if shape.auto_shape_type == MSO_SHAPE.OVAL:
            r = shape.width
            g.append(draw.Circle(shape.left+r/2, shape.top+r/2, r/2, fill='#1248ff'))
    else:
        g.append(draw.Rectangle(shape.left, shape.top, shape.width, shape.height, fill='#1248ff'))
    # d.append(draw.Rectangle(shape.left/FACTOR, pres.slide_height/FACTOR-shape.top/FACTOR, shape.width/FACTOR, shape.height/FACTOR, fill='#1248ff'))
    # g.append(draw.Text('Text at (2,1)',0.5, 0,0, text_anchor='middle', transform='scale(1,-1) translate(2,1)'))

d.append(g)
d.setRenderSize(h=640)
d.saveSvg('out.svg')

