# pptx-tablestyle
An extension module for applying style to python-pptx table

Requires python-pptx, for sure.

# Example

Apply "Medium Style 1 - Accent 3" to the created Table

```python
import tablestyle

# ...

table = shapes.add_table(rows, cols, left, top, width, height).table

tablestyle.MediumStyle1.Accent3.apply_to(table)
```

# Reference

https://github.com/scanny/python-pptx/issues/27

https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx
