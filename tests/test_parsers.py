import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import parsers


def test_parse_customer_address_inline():
    text = (
        "Заказчик: Индивидуальный предприниматель Иванов Иван Иванович "
        "Юридический адрес: г. Москва, ул. Ленина, д. 1 Почтовый адрес: г. Москва, ул. Ленина, д. 1"
    )
    data = parsers.parse_data_from_text(text)
    assert data["Данные заказчика"] == (
        "Индивидуальный предприниматель Иванов Иван Иванович\n"
        "Юридический адрес: г. Москва, ул. Ленина, д. 1"
    )
