def replace_placeholders(doc, data):
    """Replace placeholders in doc with values from data."""
    for para in doc.paragraphs:
        for run in para.runs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell, data)


def format_preview(data):
    """Return formatted string for preview widget."""
    lines = [
        data.get("Данные заказчика", ""),
        data.get("ИНН получателя", ""),
        data.get("ОГРН получателя", ""),
        f"Транспортные услуги по договору-заявке {data.get('Номер документа', '')}",
        f"По маршруту {data.get('Адрес загрузки', '')} - {data.get('Адрес разгрузки', '')}",
        f"Автомобиль: {data.get('Марка автомобиля', '')} {data.get('Номер полуприцепа', '')}",
        f"Водитель: {data.get('ФИО водителя', '')}",
        f"Дата погрузки: {data.get('Дата погрузки', '')}",
        f"Дата разгрузки: {data.get('Дата разгрузки', '')}",
        f"Стоимость перевозки: {data.get('Стоимость перевозки', '')}",
    ]
    return "\n".join(lines)
