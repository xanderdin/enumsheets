The **orig** directory contains original unprocessed sheets.
The **proc** directory contains processed enumerated sheets.


## Commands used

Numbering:

```
enumsheets.py -c config.ini orig/*dxf
```

PDF file:

```
librecad dxf2pdf -k -o result.pdf proc/*dxf
```

