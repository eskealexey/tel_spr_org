import json
import os
import re
from datetime import datetime

import pandas as pd


def xls_to_separate_json_with_stats(xls_file, output_dir='output'):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    excel_file = pd.ExcelFile(xls_file)
    sheet_names = excel_file.sheet_names

    stats = {
        "conversion_date": datetime.now().isoformat(),
        "source_file": xls_file,
        "sheets": {}
    }

    for sheet_name in sheet_names:
        df = pd.read_excel(xls_file, sheet_name=sheet_name)
        original_count = len(df)

        # Фильтрация
        if not df.empty:
            last_column = df.columns[-1]
            filtered_df = df[df[last_column].notna()]
            filtered_count = len(filtered_df)
        else:
            filtered_df = df
            filtered_count = 0

        # Создание безопасного имени файла
        safe_name = re.sub(r'[^\w\s\-_]', '', sheet_name).strip()[:100]
        json_filename = f"{safe_name}.json"
        json_path = os.path.join(output_dir, json_filename)

        # Подготовка данных
        filtered_df.reset_index(drop=True, inplace=True)
        filtered_df = filtered_df.map(lambda x: " " if pd.isna(x) else str(x))
        data_to_save = filtered_df.to_dict(orient='records')[1:]

        # Сохранение в JSON
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=2)

        # Обновление статистики
        stats["sheets"][sheet_name] = {
            "json_file": json_filename,
            "original_rows": original_count,
            "filtered_rows": filtered_count,
            "removed_rows": original_count - filtered_count,
            "columns": list(df.columns) if not df.empty else []
        }

        print(f"✓ {json_filename}: {filtered_count}/{original_count} записей")

    # Сохранение статистики
    stats_path = os.path.join(output_dir, "conversion_stats.json")
    with open(stats_path, 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    print(f"\nСтатистика сохранена в: {stats_path}")
    print(f"Всего обработано вкладок: {len(sheet_names)}")


xls_to_separate_json_with_stats("telef.xls")
