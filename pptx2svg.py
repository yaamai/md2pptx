from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE
import base64
import svgwrite

FACTOR = 720

"""

XML:
(Pdb) pp shape.element.xml
('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
 'xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" '
 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">\n'
 '  <p:nvSpPr>\n'
 '    <p:cNvPr id="3" name=""/>\n'
 '    <p:cNvSpPr/>\n'
 '    <p:nvPr/>\n'
 '  </p:nvSpPr>\n'
 '  <p:spPr>\n'
 '    <a:xfrm>\n'
 '      <a:off x="720000" y="126360"/>\n'
 '      <a:ext cx="359640" cy="360"/>\n'
 '    </a:xfrm>\n'
 '    <a:custGeom>\n'
 '      <a:avLst/>\n'
 '      <a:gdLst/>\n'
 '      <a:ahLst/>\n'
 '      <a:rect l="l" t="t" r="r" b="b"/>\n'
 '      <a:pathLst>\n'
 '        <a:path w="21600" h="21600">\n'
 '          <a:moveTo>\n'
 '            <a:pt x="0" y="0"/>\n'
 '          </a:moveTo>\n'
 '          <a:lnTo>\n'
 '            <a:pt x="21600" y="21600"/>\n'
 '          </a:lnTo>\n'
 '        </a:path>\n'
 '      </a:pathLst>\n'
 '    </a:custGeom>\n'
 '    <a:noFill/>\n'
 '    <a:ln w="18000">\n'
 '      <a:solidFill>\n'
 '        <a:srgbClr val="000000"/>\n'
 '      </a:solidFill>\n'
 '      <a:round/>\n'
 '    </a:ln>\n'
 '  </p:spPr>\n'
 '  <p:style>\n'
 '    <a:lnRef idx="0"/>\n'
 '    <a:fillRef idx="0"/>\n'
 '    <a:effectRef idx="0"/>\n'
 '    <a:fontRef idx="minor"/>\n'
 '  </p:style>\n'
 '</p:sp>\n')

libreoffice-svg:
<path xmlns="http://www.w3.org/2000/svg" fill="none" stroke="rgb(0,0,0)" stroke-width="50" stroke-linejoin="round" d="M 2000,350 L 2999,352"/>

my-impl:

"""
def _convert_shape(dwg, shape):
    print(shape.shape_type, shape, shape.left, shape.top, shape.width, shape.height)
    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        pass
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        print(shape, shape.name, "SHAPE len = {}".format(len(shape.shapes)))
        for shape in shape.shapes:
            _convert_shape(dwg, shape)
    elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
        pass
    elif shape.shape_type == MSO_SHAPE_TYPE.FREEFORM:
        for p in shape.element.spPr.custGeom.pathLst:
            path = dwg.path(fill="none", stroke="blue", stroke_width=10)
            path.translate(shape.left/FACTOR, shape.top/FACTOR)
            for command in p:
                if command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}moveTo':
                    path.push('m', command.pt.x, command.pt.y)
                elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}lnTo':
                    import pdb; pdb.set_trace()
                    path.push('l', command.pt.x/100, command.pt.y/100)
                elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}cubicBezTo':
                    path.push('c', command[0].x, command[0].y, command[1].x, command[1].y, command[2].x, command[2].y)
                else:
                    import pdb; pdb.set_trace()
            dwg.add(path)
    elif shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
        rect = dwg.rect(insert=(shape.left/FACTOR, shape.top/FACTOR),
                        size=(shape.width/FACTOR, shape.height/FACTOR), fill='gray', opacity=0.1, stroke='gray', stroke_width=10)
        dwg.add(rect)
    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        b64_data = base64.b64encode(shape.image.blob).decode()
        data_uri = 'data:{};base64,{}'.format('image/png', b64_data)
        image = dwg.image(data_uri, insert=(shape.left/FACTOR, shape.top/FACTOR),
                          size=(shape.width/FACTOR, shape.height/FACTOR))
        dwg.add(image)
    else:
        rect = dwg.rect(insert=(shape.left/FACTOR, shape.top/FACTOR),
                        size=(shape.width/FACTOR, shape.height/FACTOR), fill='blue', stroke='red', stroke_width=10)
        dwg.add(rect)


def main():
    pres = Presentation("base.pptx")
    dwg = svgwrite.Drawing(filename="out.svg", size=(1200, 1080), debug=True)
    dwg.viewbox(0, 0, pres.slide_width/FACTOR, pres.slide_height/FACTOR)

    for shape in pres.slide_master.shapes:
        _convert_shape(dwg, shape)

    dwg.save()

main()
