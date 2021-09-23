import svgwrite

slide_width = 9144000
slide_height = 5145088

dwg = svgwrite.Drawing(filename="out.svg", size=(1200, 1080), debug=True)
dwg.viewbox(0, 0, slide_width, slide_height)

dwg.add(dwg.rect(insert=(0, 0), size=(slide_width, slide_height), fill='gray', opacity=0.1, stroke='gray', stroke_width=10))
