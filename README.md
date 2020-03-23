# Excel Brush

Re-do your images on spreadsheets

### Installing Dependencies

Clone this repository, and:
```
pip3 install argparse
pip3 install pillow
pip3 install openpyxl
```

### Usage

```
python3 excel_brush.py [-h] -i IMAGE_FILE_PATH [-height HEIGHT] [-width WIDTH]
```

Inform the image path intended to be replicated and the resolution for the spreadsheet drawing:

Example:
```
python3 excel_brush.py -h
```
```
python3 excel_brush.py -i './image.jpeg'
```
```
python3 excel_brush.py -i './image.jpeg' -height 64 -width 64
```

**[GNU v3.0](https://github.com/viniciusvviterbo/ExcelBrush/blob/master/LICENSE.md)**