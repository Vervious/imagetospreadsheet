#!/usr/bin/env python
from PIL import Image
import numpy as np
from scipy import ndimage
import xlsxwriter


im = Image.open("images/pusheen.png")

print "Image size: " + str(im.size)  # Get the width and hight of the image for iterating over

SPREADSHEET_SIZE = (50, 50)
CELL_WIDTH_RATIO = 2.5  # ratio of cell width to cell height

# get numpy array
pix = np.array(im.convert("RGB"))
factx = float(SPREADSHEET_SIZE[0]) / (im.size[0] * CELL_WIDTH_RATIO)
facty = float(SPREADSHEET_SIZE[1]) / im.size[1]

print pix.shape
print factx
print facty
pix = ndimage.zoom(pix, (facty, factx, 1), order=1)
print pix.shape

newim = Image.fromarray(pix)
newim.save("images/test.png")

# generate xlsx file
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()

for x in xrange(0, newim.size[0] + 1):
    for y in xrange(0, newim.size[1] + 1):
        if x == newim.size[0] or y == newim.size[1]:
            worksheet.write_string(y, x, 'img')
            continue

        (r, g, b) = newim.getpixel((x, y))
        # generate new cell
        hexstring = '#%02x%02x%02x' % (r, g, b)
        aformat = workbook.add_format({'bg_color': hexstring})
        worksheet.write_blank(y, x, '', aformat)

workbook.close()
