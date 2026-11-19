import random
from datetime import date, timedelta
from openpyxl import Workbook

managers = ["Иванов", "Петрова", "Сидоров", "Козлова"]
products = {"Чай чёрный": 250, "Чай зелёный": 280, "Кофе зерновой": 900,
            "Кофе молотый": 750, "Сироп": 400}

wb = Workbook()
ws = wb.active
ws.append(["Дата", "Менеджер", "Товар", "Количество", "Цена", "Сумма"])

start = date(2026, 10, 1)
for _ in range(200):
    day = start + timedelta(days=random.randint(0, 29))
    product = random.choice(list(products))
    qty = random.randint(1, 20)
    price = products[product]
    ws.append([day, random.choice(managers), product, qty, price, qty * price])

wb.save("sales.xlsx")
print("Готово: sales.xlsx")
