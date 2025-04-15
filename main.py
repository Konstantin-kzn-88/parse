import win32com.client
import re
from time import sleep
import keyboard
import threading
import sys
import pythoncom


def extract_data_from_cursor_position():
    # Инициализация COM в текущем потоке
    pythoncom.CoInitialize()

    try:
        # Инициализация COM-объектов для Word и Excel
        word = win32com.client.Dispatch("Word.Application")
        excel = win32com.client.Dispatch("Excel.Application")

        # Проверка на открытые документы Word и Excel
        try:
            word_doc = word.ActiveDocument
        except Exception:
            print("Ошибка: Документ Word не открыт!")
            pythoncom.CoUninitialize()
            return

        try:
            excel_book = excel.ActiveWorkbook
            excel_sheet = excel_book.ActiveSheet
        except Exception:
            print("Ошибка: Книга Excel не открыта!")
            pythoncom.CoUninitialize()
            return

        # Получаем активное выделение в Word
        selection = word.Selection

        if selection.Information(10):  # 10 - wdWithInTable
            # Вместо работы с целой строкой, получим текст выделения или ячейки
            cell_text = selection.Text

            # Если текст ячейки слишком короткий, попробуем расширить контекст
            if len(cell_text) < 100:
                # Сначала сохраним текущую позицию
                current_range = selection.Range

                # Попробуем получить текст всей ячейки
                try:
                    cell = selection.Cells(1)
                    cell_text = cell.Range.Text
                except Exception:
                    # Если не получилось, попробуем просто увеличить выделение
                    try:
                        selection.MoveLeft(Unit=1, Count=10, Extend=1)  # Расширить выделение влево
                        selection.MoveRight(Unit=1, Count=100, Extend=1)  # Расширить выделение вправо
                        cell_text = selection.Text

                        # Вернемся к исходному выделению
                        current_range.Select()
                    except Exception:
                        print("Не удалось получить достаточный контекст.")
                        # Вернемся к исходному выделению
                        current_range.Select()

            # Если у нас всё ещё мало текста, попробуем использовать текст всей таблицы
            if len(cell_text.strip()) < 50:
                try:
                    table = selection.Tables(1)
                    row_index = 0

                    # Определим позицию ячейки
                    try:
                        row_index = selection.Information(10)  # wdStartOfRangeRowNumber
                    except Exception:
                        pass

                    # Получим данные из строки таблицы
                    if row_index > 0:
                        try:
                            # Получаем текст всей строки таблицы
                            row = table.Rows(row_index)
                            row_text = row.Range.Text
                            cell_text = row_text

                            # Если мы находимся в начале записи об оборудовании,
                            # нам может понадобиться несколько строк
                            if row_index < table.Rows.Count:
                                # Проверим еще 5 строк (или до конца таблицы)
                                for i in range(1, min(6, table.Rows.Count - row_index + 1)):
                                    try:
                                        next_row = table.Rows(row_index + i)
                                        next_row_text = next_row.Range.Text
                                        cell_text += "\n" + next_row_text
                                    except Exception:
                                        break
                        except Exception:
                            # Получаем текст всей таблицы и работаем с ним
                            table_text = table.Range.Text
                            cell_text = table_text
                except Exception:
                    print("Не удалось определить таблицу.")

            # 1. Наименование оборудования
            equipment_name = ""

            # Ищем строки до "Рег.№" - всё, что идёт до этой строки, является наименованием оборудования
            reg_num_index = -1
            lines = cell_text.split('\n')
            for i, line in enumerate(lines):
                if line.strip().startswith('Рег.№') or 'Рег.№' in line:
                    reg_num_index = i
                    break

            if reg_num_index > 0:
                # Собираем все строки до регистрационного номера
                equipment_lines = []
                for i in range(reg_num_index):
                    clean_line = re.sub(r'\|', '', lines[i]).strip()
                    if clean_line and not clean_line.startswith(('*', '+', '-')):
                        equipment_lines.append(clean_line)

                if equipment_lines:
                    # Формируем наименование оборудования из найденных строк
                    equipment_name = ' '.join(equipment_lines).strip()
                    # Удаляем "Рег" из названия, если он случайно попал туда
                    equipment_name = re.sub(r'\s+Рег\.?№?$', '', equipment_name)
                    equipment_name = re.sub(r'\s+Рег\.?№?\s*', ' ', equipment_name)

            # Если предыдущий метод не сработал, пробуем традиционные паттерны
            if not equipment_name:
                # Специфический поиск по типам оборудования
                equipment_patterns = [
                    (r'Колонна\s+([^\s|\n]+)', "Колонна"),
                    (r'Аппарат\s+([^\s|\n]+)', "Аппарат"),
                    (r'Аппарат\s+([^\s|\n]+)', "Аппарат"),
                    (r'Теплообменник\s+([^\s|\n]+)', "Теплообменник"),
                    (r'АВО\s+([^\s|\n]+)', "АВО"),
                    (r'\|\s+Колонна\s+([^\s|\n]+)', "Колонна"),
                    (r'\|\s+Теплообменник\s+([^\s|\n]+)', "Теплообменник"),
                    (r'\|\s+АВО\s+([^\s|\n]+)', "АВО")
                ]

                for pattern, equipment_type in equipment_patterns:
                    equipment_match = re.search(pattern, cell_text)
                    if equipment_match:
                        equipment_model = equipment_match.group(1).strip()
                        equipment_name = f"{equipment_type} {equipment_model}"

                        # Проверяем, есть ли дополнительные части модели (например, "АС-108 В")
                        parts_match = re.search(f"{re.escape(equipment_model)}\\s+([A-Za-zА-Яа-я0-9\\-]+)", cell_text)
                        if parts_match:
                            equipment_name += f" {parts_match.group(1)}"

                        # Удаляем "Рег" из названия, если он случайно попал туда
                        equipment_name = re.sub(r'\s+Рег\.?№?$', '', equipment_name)
                        equipment_name = re.sub(r'\s+Рег\.?№?\s*', ' ', equipment_name)
                        break

            # Если модель всё еще не найдена, пробуем найти любое обозначение оборудования
            if not equipment_name:
                # Ищем строки с наименованием оборудования (первые строки блока)
                for i, line in enumerate(lines[:3]):  # Проверяем только первые 3 строки
                    clean_line = re.sub(r'\|', '', line).strip()
                    if clean_line and not clean_line.startswith(('*', '+', '-', 'Цех', 'Секц')):
                        parts = clean_line.split()
                        if len(parts) >= 2:
                            equipment_name = clean_line
                            # Удаляем "Рег" из названия, если он случайно попал туда
                            equipment_name = re.sub(r'\s+Рег\.?№?$', '', equipment_name)
                            equipment_name = re.sub(r'\s+Рег\.?№?\s*', ' ', equipment_name)
                            break

            # 2. Опасное вещество
            dangerous_substance = ""

            # Попробуем несколько подходов для поиска опасного вещества

            # Подход 1: Прямой поиск "Опасное вещество: X"
            substance_index = -1
            for i, line in enumerate(lines):
                if 'Опасное вещество:' in line:
                    substance_index = i
                    break

            if substance_index >= 0:
                # Извлекаем опасное вещество из строки с заголовком
                line = lines[substance_index]
                parts = line.split('Опасное вещество:')
                if len(parts) > 1 and parts[1].strip():
                    dangerous_substance = parts[1].strip()

                # Проверяем следующие строки для поиска продолжения вещества
                i = substance_index + 1
                while i < len(lines) and i < substance_index + 4:  # Проверяем до 3 следующих строк
                    next_line = lines[i].strip()
                    next_line_clean = re.sub(r'[\|\s]+', ' ', next_line).strip()

                    # Останавливаемся, если встречаем строку с техническими параметрами
                    if any(x in next_line_clean for x in
                           ['Р=', 'Т=', 'V=', 'S=', 'Год']) or 'МПа' in next_line_clean or '°С' in next_line_clean:
                        break

                    # Если строка не пустая и не начинается с символов таблицы
                    if (next_line_clean and
                            not any(x in next_line_clean for x in ['Рег.№', 'Зав.№']) and
                            not re.match(r'^[+*\-]', next_line_clean)):

                        # Если у нас уже есть часть вещества, добавляем пробел
                        if dangerous_substance:
                            # Если предыдущая часть заканчивается запятой, не добавляем запятую
                            if dangerous_substance.rstrip().endswith(','):
                                dangerous_substance = dangerous_substance.rstrip() + ' ' + next_line_clean
                            else:
                                dangerous_substance = dangerous_substance + ', ' + next_line_clean
                        else:
                            dangerous_substance = next_line_clean
                    else:
                        # Если нашли строку с техническими параметрами или пустую строку, значит опасное вещество закончилось
                        if not next_line_clean or next_line_clean.isspace():
                            break
                    i += 1

                # Очистка опасного вещества от лишних символов
                dangerous_substance = re.sub(r'[\|\s]+', ' ', dangerous_substance).strip()

            # Специальные случаи для известных веществ
            if not dangerous_substance or len(
                    dangerous_substance.split()) > 6:  # Если вещество не найдено или слишком длинное (явно захватило лишнее)
                if 'Диз.топливо' in cell_text and 'водяной пар' in cell_text:
                    dangerous_substance = 'Диз.топливо, водяной пар'
                elif 'Гудрон' in cell_text and 'нефть' in cell_text:
                    dangerous_substance = 'Гудрон, нефть'
                elif 'Газойль' in cell_text and 'нефть' in cell_text:
                    dangerous_substance = 'Газойль, нефть'
                elif 'Углеводороды' in cell_text:
                    dangerous_substance = 'Углеводороды'
                elif 'Газойль' in cell_text:
                    dangerous_substance = 'Газойль'

            # Проверка и очистка результата
            # Если опасное вещество содержит технические параметры, обрезаем их
            if dangerous_substance:
                for tech_param in ['S=','V=', 'Р=', 'Т=', 'МПа', '°С', 'Год ', 'изготовления', 'эксплуатацию']:
                    if tech_param in dangerous_substance:
                        parts = dangerous_substance.split(tech_param, 1)
                        dangerous_substance = parts[0].strip()

                # Удаляем лишние символы в конце
                dangerous_substance = re.sub(r'[,\.\s]+$', '', dangerous_substance)

            # 3. Температура
            temperature = ""
            temp_patterns = [
                r'Т=\s*([\d,./]+)\s*°С',
                r'Т=\s*([\d,./]+)',
                r'температура[^=]*=\s*([\d,./]+)'
            ]

            for pattern in temp_patterns:
                temp_match = re.search(pattern, cell_text, re.IGNORECASE)
                if temp_match:
                    temperature = temp_match.group(1).strip()
                    break

            # 4. Рабочее давление
            pressure = ""
            pressure_patterns = [
                r'Р=\s*([\d,./]+)\s*МПа',
                r'Р=\s*([\d,./]+)',
                r'давление[^=]*=\s*([\d,./]+)'
            ]

            for pattern in pressure_patterns:
                pressure_match = re.search(pattern, cell_text, re.IGNORECASE)
                if pressure_match:
                    pressure = pressure_match.group(1).strip()
                    break

            # 5. Количество опасного вещества
            quantity = ""
            quantity_patterns = [
                r'Горючие\s+жидкости[\s\S]+?:\s*([\d.,]+)\s*т\.',
                r'Количество\s*=\s*([\d.,]+)\s*т\.',
                r'технологическом\s+процессе:\s*([\d.,]+)\s*т\.'
            ]

            for pattern in quantity_patterns:
                quantity_match = re.search(pattern, cell_text, re.IGNORECASE)
                if quantity_match:
                    quantity = quantity_match.group(1).strip()
                    break

            # Найдем первую пустую строку в Excel
            row = excel_sheet.Cells(excel_sheet.Rows.Count, 1).End(-4162).Row + 1  # -4162 это xlUp

            # Заполняем данные в Excel
            excel_sheet.Cells(row, 1).Value = equipment_name
            excel_sheet.Cells(row, 2).Value = dangerous_substance
            excel_sheet.Cells(row, 3).Value = temperature
            excel_sheet.Cells(row, 4).Value = pressure
            excel_sheet.Cells(row, 5).Value = quantity

            # Информируем пользователя
            print(f"Данные успешно добавлены в Excel:")
            print(f"Наименование: {equipment_name}")
            print(f"Опасное вещество: {dangerous_substance}")
            print(f"Температура: {temperature}")
            print(f"Давление: {pressure}")
            print(f"Количество опасного вещества: {quantity}")
        else:
            print("Курсор не находится в таблице Word!")

    finally:
        # Освобождаем COM-ресурсы
        pythoncom.CoUninitialize()


def on_hotkey():
    # Добавляем небольшую задержку, чтобы клавиши успели "отпуститься"
    # и не влияли на работу с Word
    sleep(0.1)

    # Запускаем функцию извлечения данных в отдельном потоке
    # чтобы избежать блокировки основного потока с обработкой клавиатуры
    extract_thread = threading.Thread(target=extract_data_from_cursor_position)
    extract_thread.daemon = True
    extract_thread.start()
    print("Горячая клавиша Alt+C нажата, запускаем извлечение данных...")


def main():
    print("Программа запущена. Нажмите Alt+C для извлечения данных из текущей позиции курсора.")
    print("Для выхода нажмите Ctrl+C в этом окне.")

    # Регистрируем горячую клавишу Alt+C
    keyboard.add_hotkey('alt+c', on_hotkey)

    try:
        # Держим программу запущенной
        keyboard.wait('ctrl+c')
    except KeyboardInterrupt:
        print("\nПрограмма завершена.")
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        # Очищаем горячие клавиши при выходе
        keyboard.unhook_all()

        # Гарантируем, что все потоки завершатся
        sys.exit(0)


if __name__ == "__main__":
    main()