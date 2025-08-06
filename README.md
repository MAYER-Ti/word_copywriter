# Word Copywriter

A simple utility to copy values from one Word document into placeholder tags of another Word document. The interface is built with PyQt5.

## Usage

Install requirements and run the application:

```bash
pip install -r requirements.txt
python main.py
```

1. Select the source document containing key-value pairs (lines formatted as `Key: Value`). PDF files are supported, including scanned documents (requires `tesseract` and `poppler` installed).
2. Select the template document with placeholders like `{{Key}}`.
3. After selecting source and template, press **Save** and choose where to store the resulting document.

For Excel templates, use the `.xlsx` format.

The resulting document will be created with the placeholders replaced by the values from the source document.
