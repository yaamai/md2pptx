from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE
import base64
import svgwrite

FACTOR = 360

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
        if shape.auto_shape_type == MSO_SHAPE.OVAL:
            r = shape.width/FACTOR
            # import pdb; pdb.set_trace()
            fill_color = 'rgb({}, {}, {})'.format(*shape.fill.fore_color.rgb)
            line_color = 'rgb({}, {}, {})'.format(*shape.line.fill.fore_color.rgb)
            line_width = shape.line.width/FACTOR
            return dwg.circle((shape.left/FACTOR+r/2, shape.top/FACTOR+r/2), r/2, fill=fill_color, stroke=line_color, stroke_width=line_width)
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        # print(shape, shape.name, shape.element.xml, "SHAPE len = {}".format(len(shape.shapes)))
        r = dwg.rect(insert=(shape.left/FACTOR, shape.top/FACTOR),
                        size=(shape.width/FACTOR, shape.height/FACTOR), fill='gray', opacity=0.1, stroke='gray', stroke_width=10)
        dwg.add(r)
        group = dwg.g()
        # group.translate(shape.left/FACTOR, shape.top/FACTOR)
        for s in shape.shapes:
            group.add(_convert_shape(dwg, s))
        return group
    elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
        pass
    elif shape.shape_type == MSO_SHAPE_TYPE.FREEFORM:
        # import pdb; pdb.set_trace()
        group = dwg.g()
        for p in shape.element.spPr.custGeom.pathLst:
            # import pdb; pdb.set_trace()
            print(shape.line.width)
            line_width = shape.line.width/FACTOR
            if shape.line.width == 0:
                # SVG's default stroke width == 1
                line_width = 30
            color = 'rgb({}, {}, {})'.format(*shape.line.color.rgb)
            path = dwg.path(fill="none", stroke=color, stroke_width=line_width)
            path.translate(shape.left/FACTOR, shape.top/FACTOR)
            path.scale((shape.width/360)/p.w, (shape.height/360)/p.h)
            # print("Scale: ", shape.left/FACTOR, shape.top/FACTOR, (shape.width/360)/p.w, (shape.height/360)/p.h, shape.width/FACTOR, shape.height/FACTOR, p.w, p.h)
            for command in p:
                if command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}moveTo':
                    path.push('M', command.pt.x, command.pt.y)
                elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}lnTo':
                    # import pdb; pdb.set_trace()
                    path.push('L', command.pt.x, command.pt.y)
                elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}cubicBezTo':
                    path.push('C', command[0].x, command[0].y, command[1].x, command[1].y, command[2].x, command[2].y)
                elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}close':
                    path.push('Z')
                else:
                    # pass
                    import pdb; pdb.set_trace()
            group.add(path)
        return group

    elif shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
        return dwg.rect(insert=(shape.left/FACTOR, shape.top/FACTOR),
                        size=(shape.width/FACTOR, shape.height/FACTOR), fill='gray', opacity=0.1, stroke='gray', stroke_width=10)

    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        b64_data = base64.b64encode(shape.image.blob).decode()
        data_uri = 'data:{};base64,{}'.format('image/png', b64_data)
        return dwg.image(data_uri, insert=(shape.left/FACTOR, shape.top/FACTOR),
                         size=(shape.width/FACTOR, shape.height/FACTOR))

    return dwg.rect(insert=(shape.left/FACTOR, shape.top/FACTOR),
                    size=(shape.width/FACTOR, shape.height/FACTOR), fill='blue', stroke='red', stroke_width=10)


def main():
    pres = Presentation("base.pptx")
    dwg = svgwrite.Drawing(filename="out.svg", size=(1200, 1080), debug=True)
    dwg.viewbox(0, 0, pres.slide_width/FACTOR, pres.slide_height/FACTOR)

    for shape in pres.slide_master.shapes:
        dwg.add(_convert_shape(dwg, shape))

    dwg.save()

main()
