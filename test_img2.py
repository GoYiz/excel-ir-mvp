from openpyxl import load_workbook
wb=load_workbook('/var/minis/workspace/excel_ir_mvp/tmp_img.xlsx')
ws=wb.active
im=ws._images[0]
print(type(im), im.width, im.height, type(im.anchor))
a=im.anchor
print(a.__dict__)
print('from', a._from.__dict__)
print('ext', a.ext.__dict__)
print('data', len(im._data()), im.path)
