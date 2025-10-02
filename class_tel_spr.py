import json
import os

import pandas as pd


class ParserXls:
    def __init__(self, filexls):
        if not os.path.exists(filexls):
            raise FileNotFoundError(f"Файл {filexls} не найден")
        self.filexls = filexls
        self.all_data = {}


    def __str__(self):
        return self.filexls

    def parser(self, output_dir="JSON", output_file="all_sheets.json"):
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        excel_file = pd.ExcelFile(self.filexls)
        sheet_names = excel_file.sheet_names

        all_data = {}

        for sheet_name in sheet_names:
            df = pd.read_excel(self.filexls, sheet_name=sheet_name)

            # Замена NaN на пустые строки для корректного JSON
            df = df.fillna('')

            # Сохраняем данные листа как список словарей
            all_data[sheet_name] = df.to_dict(orient='records')

        json_path = os.path.join(output_dir, output_file)

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)

        self.all_data = all_data

        return self.all_data


class SprOsfr(ParserXls):
    """
    Класс для обработки Excel-файла с выделением нужных листов.
    """
    target_sheets = ["ОСФР"]  # Листы, которые нужно обработать
    target_map = {
        'Unnamed: 0': 'Городской номер',
        'Unnamed: 1': 'Кор. тел.',
        'Unnamed: 2': '№ комн.',
        'Unnamed: 3': 'ФАМИЛИЯ',
        'Unnamed: 4': 'ИМЯ',
        'Unnamed: 5': 'ОТЧЕСТВО',
        'Unnamed: 6': 'ДОЛЖНОСТЬ',
        'Unnamed: 7': 'Отдел',
        'Unnamed: 8': 'Место расположения'
    }

    def __init__(self, filexls):
        super().__init__(filexls)
        self.osfr = None

    def obrabotka_osfr(self):
        list_strok = []
        data = self.parser()
        for sheet_name in data.keys():
            if sheet_name in self.target_sheets:
                # print(sheet_name)
                for row in data[sheet_name]:
                    if row.get("Unnamed: 8", '') != '':
                        # Создаем новую строку с замененными ключами
                        new_row = {}
                        for old_key, value in row.items():
                            # Если ключ есть в target_map, используем новое название, иначе оставляем старый
                            new_key = self.target_map.get(old_key, old_key)
                            new_row[new_key] = value
                        list_strok.append(new_row)
        if len(list_strok) > 1:
            list_strok_ = list_strok[1:]
        else:
            list_strok_ = list_strok

        return list_strok_

    def save_to_json(self, output_dir="JSON", output_file='osfr.json'):
        """
        Сохраняет результат obrabotka_osfr в JSON файл
        """
        json_path = os.path.join(output_dir, output_file)
        result = self.obrabotka_osfr()
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)
        print(f"Данные успешно записаны в {json_path}")


class SprKs(ParserXls):
    target_sheets = ["Клиентские службы",]  # Листы, которые нужно обработать
    target_map = {
        "Unnamed: 0": "кспд",
        "Unnamed: 1": "город",
        "Unnamed: 2": "Фамилия",
        "Unnamed: 3": "Имя",
        "Unnamed: 4": "Отчество",
        "Unnamed: 5": "Должность",
        "Unnamed: 6": "Место расположения",
        # "Unnamed: 7": "Отдел"
    }

    def __init__(self, filexls):
        super().__init__(filexls)
        self.ks = None

    def obrabotka_ks(self):
        list_strok = []
        otdel = None
        required_chars = "Клиентская служба"
        data = self.parser()

        for sheet_name in data.keys():
            if sheet_name in self.target_sheets:
                for row in data[sheet_name]:
                    value = row.get("Unnamed: 0", '')
                    if required_chars in str(value):
                        otdel = value

                    if row.get("Unnamed: 6", '') != '':
                        # Создаем новую строку с замененными ключами
                        new_row = {}
                        new_row['отдел'] = otdel
                        for old_key, value in row.items():
                            new_key = self.target_map.get(old_key, old_key)
                            new_row[new_key] = value

                        list_strok.append(new_row)
        if len(list_strok) > 1:
            list_strok_ = list_strok[1:]
        else:
            list_strok_ = list_strok

        return list_strok_

    def save_to_json(self, output_dir="JSON", output_file='ks.json'):
        """
        Сохраняет результат obrabotka_osfr в JSON файл
        """
        json_path = os.path.join(output_dir, output_file)
        result = self.obrabotka_ks()
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)
        print(f"Данные успешно записаны в {json_path}")


# if __name__ == '__main__':
#     ks = SprKs("telef.xls")
#     lst_ks = ks.obrabotka_ks()
#     ks.save_to_json()
if __name__ == '__main__':
    try:
        # Обработка Клиентских служб
        ks = SprKs("telef.xls")
        lst_ks = ks.obrabotka_ks()
        ks.save_to_json()

        # Обработка ОСФР
        osfr = SprOsfr("telef.xls")
        lst_osfr = osfr.obrabotka_osfr()
        osfr.save_to_json(output_file='osfr.json')

        print("Обработка завершена успешно!")

    except FileNotFoundError as e:
        print(f"Ошибка: {e}")
    except Exception as e:
        print(f"Неожиданная ошибка: {e}")