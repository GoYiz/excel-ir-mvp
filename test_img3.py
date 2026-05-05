from openpyxl import load_workbook
wb=load_workbook('/var/minis/workspace/excel_ir_mvp/tmp_img.xlsx')
ws=wb.active
im=ws._images[0]
print(type(im).__name__, im.width, im.height, type(im.anchor).__name__)
a=im.anchor
print('from row col', a._from.row, a._from.col, a._from.rowOff, a._from.colOff)
print('ext', a.ext.cx, a.ext.cy)
print('data', len(im._data()), im.path, im.format)
