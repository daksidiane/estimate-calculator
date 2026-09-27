from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta

def create_database():
    wb = Workbook()
    
    # Стили
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # ========== ЛИСТ 1: Типы конструкций ==========
    ws_constructions = wb.active
    ws_constructions.title = "Constructions"
    
    constructions_headers = ["ID", "Code", "Name", "Description", "BasePrice", "IsActive"]
    constructions_data = [
        [1, "WINDOW", "Окно", "Стандартное окно", 0, True],
        [2, "DOOR", "Дверь", "Входная/межкомнатная дверь", 0, True],
        [3, "FACADE", "Фасад", "Фасадная система", 0, True],
        [4, "PARTITION", "Перегородка", "Внутренняя перегородка", 0, True],
        [5, "BALCONY", "Балкон", "Балконное остекление", 0, True],
        [6, "SHOWER", "Душевая кабина", "Стеклянная душевая кабина", 0, True],
    ]
    
    ws_constructions.append(constructions_headers)
    for row in constructions_data:
        ws_constructions.append(row)
    
    # ========== ЛИСТ 2: Компоненты (Стекло, Профиль, Фурнитура) ==========
    ws_components = wb.create_sheet("Components")
    
    components_headers = [
        "ID", "Category", "Code", "Name", "Description", 
        "Unit", "Price", "MinPrice", "MaxPrice", "IsActive", "CreatedDate"
    ]
    components_data = [
        # СТЕКЛО
        [1, "GLASS", "GL-001", "Стекло 4мм", "Прозрачное стекло 4мм", "м²", 850, 850, 850, True, "2026-01-01"],
        [2, "GLASS", "GL-002", "Стекло 6мм", "Прозрачное стекло 6мм", "м²", 1200, 1200, 1200, True, "2026-01-01"],
        [3, "GLASS", "GL-003", "Стекло 8мм", "Прозрачное стекло 8мм", "м²", 1650, 1650, 1650, True, "2026-01-01"],
        [4, "GLASS", "GL-004", "Стекло 10мм", "Прозрачное стекло 10мм", "м²", 2100, 2100, 2100, True, "2026-01-01"],
        [5, "GLASS", "GL-005", "Тонированное 4мм", "Тонированное стекло 4мм", "м²", 1100, 1100, 1100, True, "2026-01-01"],
        [6, "GLASS", "GL-006", "Закаленное 6мм", "Закаленное стекло 6мм", "м²", 1800, 1800, 1800, True, "2026-01-01"],
        [7, "GLASS", "GL-007", "Закаленное 8мм", "Закаленное стекло 8мм", "м²", 2400, 2400, 2400, True, "2026-01-01"],
        [8, "GLASS", "GL-008", "Триплекс 4.4.2", "Многослойное стекло", "м²", 2800, 2800, 2800, True, "2026-01-01"],
        [9, "GLASS", "GL-009", "Триплекс 6.6.2", "Многослойное стекло усиленное", "м²", 3500, 3500, 3500, True, "2026-01-01"],
        [10, "GLASS", "GL-010", "Энергосберегающее", "i-стекло с покрытием", "м²", 1950, 1950, 1950, True, "2026-01-01"],
        
        # ПРОФИЛЬ АЛЮМИНИЕВЫЙ
        [11, "PROFILE", "PR-001", "Alumil A150", "Алюминиевый профиль холодный", "пог.м", 450, 450, 450, True, "2026-01-01"],
        [12, "PROFILE", "PR-002", "Alumil A250", "Алюминиевый профиль теплый", "пог.м", 680, 680, 680, True, "2026-01-01"],
        [13, "PROFILE", "PR-003", "Schuco AWS", "Система Schuco AWS", "пог.м", 950, 950, 950, True, "2026-01-01"],
        [14, "PROFILE", "PR-004", "Reynaers CP155", "Фасадная система", "пог.м", 1200, 1200, 1200, True, "2026-01-01"],
        [15, "PROFILE", "PR-005", "Provedal P3", "Раздвижная система", "пог.м", 520, 520, 520, True, "2026-01-01"],
        [16, "PROFILE", "PR-006", "Alutech F50", "Фасадная система", "пог.м", 750, 750, 750, True, "2026-01-01"],
        
        # ФУРНИТУРА
        [17, "HARDWARE", "HW-001", "Петля стандарт", "Петля для окон/дверей", "шт", 250, 250, 250, True, "2026-01-01"],
        [18, "HARDWARE", "HW-002", "Петля усиленная", "Усиленная петля", "шт", 450, 450, 450, True, "2026-01-01"],
        [19, "HARDWARE", "HW-003", "Ручка поворотная", "Ручка оконная", "шт", 350, 350, 350, True, "2026-01-01"],
        [20, "HARDWARE", "HW-004", "Ручка-кнопка", "Ручка для дверей", "шт", 280, 280, 280, True, "2026-01-01"],
        [21, "HARDWARE", "HW-005", "Замок многозапорный", "Многоточечный замок", "шт", 1850, 1850, 1850, True, "2026-01-01"],
        [22, "HARDWARE", "HW-006", "Защелка", "Дверная защелка", "шт", 180, 180, 180, True, "2026-01-01"],
        [23, "HARDWARE", "HW-007", "Доводчик", "Дверной доводчик", "шт", 2400, 2400, 2400, True, "2026-01-01"],
        [24, "HARDWARE", "HW-008", "Уплотнитель", "Резиновый уплотнитель", "пог.м", 45, 45, 45, True, "2026-01-01"],
        [25, "HARDWARE", "HW-009", "Уголок соединительный", "Уголок 90 градусов", "шт", 120, 120, 120, True, "2026-01-01"],
        [26, "HARDWARE", "HW-010", "Ролики раздвижные", "Комплект роликов", "комплект", 650, 650, 650, True, "2026-01-01"],
    ]
    
    ws_components.append(components_headers)
    for row in components_data:
        ws_components.append(row)
    
    # ========== ЛИСТ 3: Дополнительные работы ==========
    ws_works = wb.create_sheet("AdditionalWorks")
    
    works_headers = [
        "ID", "Code", "Name", "Description", "Unit", "Price", "CalculationType", "IsActive"
    ]
    works_data = [
        [1, "AW-001", "Герметизация", "Герметизация швов силиконом", "пог.м", 150, "PER_METER", True],
        [2, "AW-002", "Установка отлива", "Монтаж водоотлива", "шт", 450, "FIXED", True],
        [3, "AW-003", "Установка подоконника", "Монтаж подоконника", "шт", 800, "FIXED", True],
        [4, "AW-004", "Откосы пластиковые", "Установка пластиковых откосов", "пог.м", 550, "PER_METER", True],
        [5, "AW-005", "Откосы гипсокартон", "Установка ГКЛ откосов", "пог.м", 650, "PER_METER", True],
        [6, "AW-006", "Москитная сетка", "Изготовление и установка", "шт", 1200, "FIXED", True],
        [7, "AW-007", "Ламинация профиля", "Цветная ламинация", "м²", 850, "PER_SQM", True],
        [8, "AW-008", "Покраска профиля", "Порошковая покраска", "м²", 950, "PER_SQM", True],
        [9, "AW-009", "Сверление отверстий", "Дополнительные отверстия", "шт", 200, "FIXED", True],
        [10, "AW-010", "Усиление конструкции", "Дополнительное усиление", "шт", 1500, "FIXED", True],
    ]
    
    ws_works.append(works_headers)
    for row in works_data:
        ws_works.append(row)
    
    # ========== ЛИСТ 4: Услуги (Монтаж, Замер, Доставка) ==========
    ws_services = wb.create_sheet("Services")
    
    services_headers = [
        "ID", "Code", "Name", "Description", "Unit", "Price", "MinPrice", "CalculationType", "IsActive"
    ]
    services_data = [
        [1, "SRV-001", "Замер", "Выезд замерщика", "выезд", 0, 1000, "FIXED_MIN", True],
        [2, "SRV-002", "Доставка", "Доставка материалов", "рейс", 1500, 1500, "FIXED", True],
        [3, "SRV-003", "Доставка срочная", "Срочная доставка день в день", "рейс", 2500, 2500, "FIXED", True],
        [4, "SRV-004", "Подъем на этаж", "Подъем материалов", "этаж", 200, 200, "PER_FLOOR", True],
        [5, "SRV-005", "Монтаж окна", "Установка окна", "шт", 2500, 2500, "FIXED", True],
        [6, "SRV-006", "Монтаж двери", "Установка двери", "шт", 3500, 3500, "FIXED", True],
        [7, "SRV-007", "Монтаж фасада", "Установка фасадной системы", "м²", 1800, 1800, "PER_SQM", True],
        [8, "SRV-008", "Монтаж перегородки", "Установка перегородки", "м²", 1500, 1500, "PER_SQM", True],
        [9, "SRV-009", "Демонтаж", "Демонтаж старых конструкций", "шт", 800, 800, "FIXED", True],
        [10, "SRV-010", "Консультация", "Техническая консультация", "час", 1000, 1000, "PER_HOUR", True],
    ]
    
    ws_services.append(services_headers)
    for row in services_data:
        ws_services.append(row)
    
    # ========== ЛИСТ 5: Скидки ==========
    ws_discounts = wb.create_sheet("Discounts")
    
    discounts_headers = [
        "ID", "Code", "Name", "Description", "DiscountPercent", "MinAmount", "MaxAmount", 
        "ApplicableTo", "IsActive", "ValidFrom", "ValidTo"
    ]
    discounts_data = [
        [1, "DISC-001", "Опт 10%", "Скидка при заказе от 100000", 10, 100000, None, "ALL", True, "2026-01-01", "2026-12-31"],
        [2, "DISC-002", "Опт 15%", "Скидка при заказе от 200000", 15, 200000, None, "ALL", True, "2026-01-01", "2026-12-31"],
        [3, "DISC-003", "Опт 20%", "Скидка при заказе от 500000", 20, 500000, None, "ALL", True, "2026-01-01", "2026-12-31"],
        [4, "DISC-004", "Постоянный клиент", "Скидка постоянным клиентам", 5, 0, None, "ALL", True, "2026-01-01", "2026-12-31"],
        [5, "DISC-005", "Акция весна", "Сезонная акция", 7, 50000, 15, "WINDOW", True, "2026-03-01", "2026-05-31"],
        [6, "DISC-006", "Комплексный заказ", "При заказе 3+ конструкций", 12, 0, None, "ALL", True, "2026-01-01", "2026-12-31"],
    ]
    
    ws_discounts.append(discounts_headers)
    for row in discounts_data:
        ws_discounts.append(row)
    
    # ========== ЛИСТ 6: Коэффициенты расчета ==========
    ws_coeffs = wb.create_sheet("CalculationCoefficients")
    
    coeffs_headers = [
        "ID", "Code", "Name", "Description", "Coefficient", "ApplicableTo", "IsActive"
    ]
    coeffs_data = [
        [1, "COEF-001", "Срочность x1.5", "Срочное изготовление", 1.5, "ALL", True],
        [2, "COEF-002", "Сложная форма", "Нестандартная форма", 1.3, "ALL", True],
        [3, "COEF-003", "Большой размер", "Конструкция > 5м²", 1.15, "ALL", True],
        [4, "COEF-004", "Малый размер", "Конструкция < 0.5м²", 1.4, "ALL", True],
        [5, "COEF-005", "Высотные работы", "Работа на высоте > 10м", 1.25, "MOUNTING", True],
        [6, "COEF-006", "Зимний коэффициент", "Работа в зимний период", 1.1, "MOUNTING", True],
    ]
    
    ws_coeffs.append(coeffs_headers)
    for row in coeffs_data:
        ws_coeffs.append(row)
    
    # ========== ЛИСТ 7: Формулы расчета ==========
    ws_formulas = wb.create_sheet("Formulas")
    
    formulas_headers = [
        "ID", "ConstructionType", "FormulaName", "Formula", "Description"
    ]
    formulas_data = [
        [1, "WINDOW", "AREA", "WIDTH * HEIGHT", "Площадь окна"],
        [2, "WINDOW", "GLASS_COST", "AREA * GLASS_PRICE", "Стоимость стекла"],
        [3, "WINDOW", "PROFILE_COST", "PERIMETER * PROFILE_PRICE", "Стоимость профиля"],
        [4, "WINDOW", "TOTAL", "GLASS_COST + PROFILE_COST + HARDWARE_COST", "Итого по окну"],
        [5, "DOOR", "AREA", "WIDTH * HEIGHT", "Площадь двери"],
        [6, "DOOR", "TOTAL", "AREA * PRICE_PER_SQM + HARDWARE_COST", "Итого по двери"],
        [7, "FACADE", "AREA", "WIDTH * HEIGHT", "Площадь фасада"],
        [8, "FACADE", "TOTAL", "AREA * SYSTEM_PRICE + GLASS_COST + MOUNTING_COST", "Итого по фасаду"],
    ]
    
    ws_formulas.append(formulas_headers)
    for row in formulas_data:
        ws_formulas.append(row)
    
    # ========== ЛИСТ 8: Единицы измерения ==========
    ws_units = wb.create_sheet("Units")
    
    units_headers = ["ID", "Code", "Name", "Description"]
    units_data = [
        [1, "PCS", "Штука", "Количество штук"],
        [2, "SQM", "Квадратный метр", "Площадь в м²"],
        [3, "LM", "Погонный метр", "Длина в погонных метрах"],
        [4, "M", "Метр", "Длина в метрах"],
        [5, "SET", "Комплект", "Комплект изделий"],
        [6, "HOUR", "Час", "Время в часах"],
        [7, "FLOOR", "Этаж", "Количество этажей"],
        [8, "FIXED", "Фиксировано", "Фиксированная цена"],
    ]
    
    ws_units.append(units_headers)
    for row in units_data:
        ws_units.append(row)
    
    # ========== ФОРМАТИРОВАНИЕ ==========
    def format_sheet(ws):
        # Автоширина колонок
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Форматирование заголовков
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # Форматирование данных
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center')
    
    # Применяем форматирование ко всем листам
    for sheet_name in wb.sheetnames:
        format_sheet(wb[sheet_name])
    
    # ========== СОХРАНЕНИЕ ==========
    filename = "calculator_database.xlsx"
    wb.save(filename)
    print(f"✅ База данных успешно создана: {filename}")
    print(f"📊 Создано листов: {len(wb.sheetnames)}")
    print(f"📋 Листы: {', '.join(wb.sheetnames)}")
    
    return filename

if __name__ == "__main__":
    create_database()