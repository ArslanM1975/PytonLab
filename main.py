import pandas as pd
import os
import re

# Параметры и список файлов
input_files = [
    'test/file 22.07.2023.xlsx', 'test/file 23.07.2023.xlsx', 'test/file 24.07.2023.xlsx',
    'test/file 25.07.2023.xlsx', 'test/file 26.07.2023.xlsx', 'test/file 27.07.2023.xlsx',
    'test/file 28.07.2023.xlsx'
]  # Список файлов

# Создаем пустой DataFrame для объединения данных
all_data = []

# Считываем данные из каждого файла
for file in input_files:
    if os.path.isfile(file):
        print(f"Чтение файла: {file}")
        try:
            # Читаем данные из файла, без заголовков
            df = pd.read_excel(file, header=None)
            df.columns = ['ФИО', 'Статус', 'Сведения об отсутствующих', 'Start', 'End']  # Установим имена столбцам

            # Извлекаем дату из имени файла
            date_str = re.search(r'(\d{2}\.\d{2}\.\d{4})', file)
            if date_str:
                date = date_str.group(0)
                # Добавляем новую колонку "Дата"
                df['Дата'] = date

            all_data.append(df[['ФИО', 'Статус', 'Сведения об отсутствующих', 'Дата']])  # Сохраняем нужные столбцы
            print("Файл прочитан успешно.")
        except Exception as e:
            print(f"Ошибка при чтении файла {file}: {e}")
    else:
        print(f"Файл не найден: {file}")

# Объединяем все данные в один DataFrame
if all_data:
    all_data_df = pd.concat(all_data, ignore_index=True)

    # Создаем столбец с комбинированной информацией о статусе и примечаниях
    all_data_df['Статус с примечанием'] = all_data_df.apply(
        lambda row: f"{row['Статус']} ({row['Сведения об отсутствующих']})" if pd.notnull(row['Сведения об отсутствующих']) else row['Статус'],
        axis=1
    )

    # Создаем сводную таблицу
    pivot_table = all_data_df.pivot_table(index='ФИО', columns='Дата', values='Статус с примечанием', aggfunc='first')

    # Заполняем пропуски
    pivot_table.fillna('Нет данных', inplace=True)

    # Выводим таблицу
    print(pivot_table)

    # Сохраняем в Excel
    output_file = 'status_table.xlsx'
    pivot_table.to_excel(output_file, sheet_name='Статусы', index=True)

    print(f"Таблица сохранена в файл: {output_file}")
else:
    print("Не было считано ни одного файла.")
