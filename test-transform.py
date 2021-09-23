import svgwrite

slide_width = 9144000
slide_height = 5145088

grp_x = 7607300
grp_y = 114300
grp_w = 1417638
grp_h = 792163

grp_ch_x = 4792
grp_ch_y = 72
grp_ch_w = 893
grp_ch_h = 499

sp_x = 5232
sp_y = 157
sp_w = 113
sp_h = 88

path_w = 2365
path_h = 1825
path_cmd = "M 961 764 C 875 688 736 643 599 642 C 270 641 2 900 1 1235 C 0 1563 270 1824 599 1825 C 783 1825 943 1751 1053 1619 C 1053 1293 1053 1293 1053 1293 C 994 1394 877 1524 771 1568 C 717 1590 662 1604 599 1604 C 394 1604 225 1445 225 1235 C 225 1041 380 862 599 863 C 701 863 793 905 861 972 C 931 1040 1041 1183 1093 1239 C 1227 1383 1419 1473 1630 1473 C 2036 1474 2365 1146 2365 741 C 2365 336 2035 0 1630 0 C 1395 0 1187 122 1053 295 C 1053 735 1053 735 1053 735 C 1155 455 1340 232 1630 232 C 1909 232 2135 463 2134 741 C 2134 1020 1909 1246 1630 1245 C 1506 1244 1391 1199 1303 1125 C 1194 1040 1074 860 961 764"

#      <a:chOff x="4792" y="72"/>
#      <a:chExt cx="893" cy="499"/>

view_height = 540
dwg = svgwrite.Drawing(filename="out.svg", size=(slide_width*view_height/slide_height, view_height), debug=True)
dwg.viewbox(0, 0, slide_width, slide_height)

# slide rect
dwg.add(dwg.rect(insert=(0, 0), size=(slide_width, slide_height), fill='gray', opacity=0.1, stroke='gray', stroke_width=10))

# logo rect
dwg.add(dwg.rect(insert=(grp_x, grp_y), size=(grp_w, grp_h), fill='gray', opacity=0.1, stroke='gray', stroke_width=10))

# infinity mark rect
f_w = grp_w/grp_ch_w# * grp_ch_w/sp_w
f_h = grp_h/grp_ch_h# * grp_ch_h/sp_h
# dwg.add(dwg.rect(insert=(grp_x+(sp_x-grp_ch_x)*f_w, grp_y+(sp_y-grp_ch_y)*f_h), size=(sp_w*f_w, sp_h*f_h), fill='gray', opacity=0.1, stroke='gray', stroke_width=10))
# dwg.add(dwg.rect(insert=((grp_x/grp_ch_x)*sp_x, (grp_y/grp_ch_y)*sp_y), size=(sp_w*f_w, sp_h*f_h), fill='gray', opacity=0.1, stroke='gray', stroke_width=10))
# equals. print(f_w, f_h, (grp_x/grp_ch_x), (grp_y/grp_ch_y))
dwg.add(dwg.rect(insert=(grp_x+sp_x*f_w-grp_ch_x*f_w, grp_y+sp_y*f_h-grp_ch_y*f_h), size=(sp_w*f_w, sp_h*f_h), fill='gray', opacity=0.1, stroke='gray', stroke_width=10))

g = dwg.g()
g.translate(grp_x+sp_x*f_w-grp_ch_x*f_w, grp_y+sp_y*f_h-grp_ch_y*f_h)

p = dwg.path(path_cmd, fill="red")
p.scale((sp_w*f_w)/path_w, (sp_h*f_h)/path_h)
# p.translate(grp_x+sp_x*f_w-grp_ch_x*f_w, grp_y+sp_y*f_h-grp_ch_y*f_h)
# p.translate(grp_x+sp_x*f_w-grp_ch_x*f_w,0 )
# maybe svg's translate are incorrect? or use g's transform insteadd of path's transform
g.add(p)
dwg.add(g)

dwg.save()

