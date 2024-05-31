import os
import random
from docxtpl import DocxTemplate

# Шаблон документа
template_file = "tmp.docx"

# Список данных для счетов
invoices = [
    {
        "company": random.choice(["Анталья Продуктс", "Юнилевер", "Грифиндор"]),
        "check_number": random.randint(1000, 9999),
        "day": f" {random.randint(1, 28)} .",
        "month": f" {random.randint(1, 12)} . ",
        "year": random.randint(21, 24),
        "seller": random.choice(["ООО ДНС", "ОАО РОСТРАНС", "ЗАО МАКС", "РОСИНГОС", "АО ЭЛТЕХНО"]),
        "address": random.choice(["ул. Пушкина, д. Калатушкина", "пр. Мира, ул. Кефира", "ул. Профсоюзная,15"]),
        "ORGN": f"{random.randint(1000000000, 9999999999)}",
        "products": [
            {
                "title": f"Товар {random.randint(1, 50)}",
                "code": f"ART-{random.randint(100, 999)}",
                "unit": random.choice(["шт", "кг", "л", "м"]),
                "amount": random.randint(1, 100),
                "price": round(random.uniform(10.0, 1000.0), 2),
                "sum": round(random.uniform(10.0, 1000.0), 2)
            } for _ in range(random.randint(3, 6))
        ],
        "general_sum": round(sum(product["sum"] for product in [
            {
                "title": f"Товар {random.randint(1, 50)}",
                "code": f"ART-{random.randint(100, 999)}",
                "unit": random.choice(["шт", "кг", "л", "м"]),
                "amount": random.randint(1, 100),
                "price": round(random.uniform(10.0, 1000.0), 2),
                "sum": round(random.uniform(10.0, 1000.0), 2)
            } for _ in range(random.randint(3, 6))
        ]), 2)
    } for _ in range(15)
]

# Создание директории для сохранения счетов
output_dir = "invoices"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Создание и сохранение счетов
for i, invoice in enumerate(invoices):
    doc = DocxTemplate(template_file)
    doc.render(invoice)
    output_file = os.path.join(output_dir, f"invoice_{i+1}.docx")
    doc.save(output_file)
    print(f"Счет {i+1} сохранен как {output_file}")
