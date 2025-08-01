# Word Copywriter

A simple utility to copy values from one Word document into placeholder tags of another Word document. The interface is built with PyQt5.

## Usage

Install requirements and run the application:

```bash
pip install -r requirements.txt
python word_copywriter.py
```

1. Select the source document containing key-value pairs (lines formatted as `Key: Value`).
2. Select the template document with placeholders like `{{Key}}`.
3. Choose where to save the resulting document and press **Copy**.

The resulting document will be created with the placeholders replaced by the values from the source document.
