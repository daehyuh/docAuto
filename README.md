# docAuto
docAuto

```shell
pip install python-docx
```

```python
import docx
import glob
from docx.shared import Cm, Inches, Mm

folder = "" #image folder

img_list = glob.glob(folder+'/*.png')

list(set(img_list))
img_list.sort()
print(img_list, len(img_list))

document = docx.Document()

tbl = document.add_table(rows=0, cols=1)
tbl.style = document.styles['Table Grid']

for i in img_list:
    print(i)
    row_cells = tbl.add_row().cells
    paragraph = row_cells[0].paragraphs[0]
    run = paragraph.add_run()
    run.add_picture(i, width= Mm(54.57), height= Mm(30.93))

document.save("output.docx") #save file
print('end')
```
