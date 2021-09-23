from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL
import base64
import svgwrite


def _calc_slide_size(pres, width=None, height=720):
    if height:
        return (pres.slide_width*height/pres.slide_height, height)
    if width:
        return (width, pres.slide_height*width/pres.slide_width)
    raise Exception()


def _get_svg_color_info(shape):
    fill = "none"
    if shape.fill.type == MSO_FILL.SOLID:
        fill = 'rgb({}, {}, {})'.format(*shape.fill.fore_color.rgb)
    stroke_color = ""
    stroke_width = shape.line.width
    if shape.line.width != 0:
        stroke_color = 'rgb({}, {}, {})'.format(*shape.line.color.rgb)

    return fill, stroke_color, stroke_width

def _convert_shape_auto_shape(dwg, shape, parent):
    if shape.auto_shape_type == MSO_SHAPE.OVAL:
        r = shape.width
        fill_color, stroke_color, stroke_width = _get_svg_color_info(shape)
        return dwg.circle((shape.left+r/2, shape.top+r/2), r/2, fill=fill_color, stroke=stroke_color, stroke_width=stroke_width)

    if shape.auto_shape_type == MSO_SHAPE.RECTANGLE:
        fill, stroke_color, stroke_width = _get_svg_color_info(shape)
        return dwg.rect(insert=(shape.left, shape.top),
                        size=(shape.width, shape.height), fill=fill, stroke=stroke_color, stroke_width=stroke_width)


def _convert_shape_group(dwg, shape, parent):
    # print(shape, shape.name, shape.element.xml, "SHAPE len = {}".format(len(shape.shapes)))
    # r = dwg.rect(insert=(shape.left, shape.top),
    #                 size=(shape.width, shape.height), fill='blue', opacity=0.1, stroke='gray', stroke_width=10)
    # dwg.add(r)

    group = dwg.g()
    group.translate(shape.left, shape.top)
    for s in shape.shapes:
        elem = _convert_shape(dwg, s, shape)
        if not elem:
            print("skipped shape: ", s)
        if elem:
            group.add(elem)
    return group


def _convert_shape_line(dwg, shape, parent):
    pass


def _convert_shape_freeform(dwg, shape, parent):
    ff_group = dwg.g()
    for p in shape.element.spPr.custGeom.pathLst:
        # wrap path to group element due to translate coords incorrectly if using shape's translate func
        g = dwg.g()

        f_w = parent.width/parent.element.xfrm.chExt.cx
        f_h = parent.height/parent.element.xfrm.chExt.cy
        g.translate(shape.left*f_w-parent.element.xfrm.chOff.x*f_w, shape.top*f_h-parent.element.xfrm.chOff.y*f_h)

        fill, stroke_color, stroke_width = _get_svg_color_info(shape)
        path = dwg.path(fill=fill, stroke=stroke_color, stroke_width=stroke_width)
        path.scale((shape.width*f_w)/p.w, (shape.height*f_h)/p.h)

        for command in p:
            if command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}moveTo':
                path.push('M', command.pt.x, command.pt.y)
            elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}lnTo':
                path.push('L', command.pt.x, command.pt.y)
            elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}cubicBezTo':
                path.push('C', command[0].x, command[0].y, command[1].x, command[1].y, command[2].x, command[2].y)
            elif command.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}close':
                path.push('Z')
            else:
                import pdb; pdb.set_trace()
        g.add(path)
        ff_group.add(g)
    return ff_group


def _convert_shape_placeholder(dwg, shape, parent):
    return dwg.rect(insert=(shape.left, shape.top),
                    size=(shape.width, shape.height), fill='gray', opacity=0.1, stroke='gray', stroke_width=10)


def _convert_shape_picture(dwg, shape, parent):
    b64_data = base64.b64encode(shape.image.blob).decode()
    data_uri = 'data:{};base64,{}'.format('image/png', b64_data)
    return dwg.image(data_uri, insert=(shape.left, shape.top),
                     size=(shape.width, shape.height))

SHAPE_TYPE_FUNCS = {
    MSO_SHAPE_TYPE.AUTO_SHAPE: _convert_shape_auto_shape,
    MSO_SHAPE_TYPE.GROUP: _convert_shape_group,
    MSO_SHAPE_TYPE.LINE: _convert_shape_line,
    MSO_SHAPE_TYPE.FREEFORM: _convert_shape_freeform,
    MSO_SHAPE_TYPE.PLACEHOLDER: _convert_shape_placeholder,
    MSO_SHAPE_TYPE.PICTURE: _convert_shape_picture,
}


def _convert_shape(dwg, shape, parent):
    if shape.shape_type not in SHAPE_TYPE_FUNCS:
        return dwg.rect(insert=(shape.left, shape.top),
                            size=(shape.width, shape.height), fill='blue', stroke='red', stroke_width=10)
    return SHAPE_TYPE_FUNCS[shape.shape_type](dwg, shape, parent)


def main():
    pres = Presentation("fj2.pptx")
    dwg = svgwrite.Drawing(filename="out.svg", size=_calc_slide_size(pres, height=720), debug=True)
    dwg.viewbox(0, 0, pres.slide_width, pres.slide_height)

    for shape in pres.slide_master.shapes:
        elem = _convert_shape(dwg, shape, None)
        if not elem:
            print("skipped shape: ", shape)
        if elem:
            dwg.add(elem)

    dwg.save()

main()
