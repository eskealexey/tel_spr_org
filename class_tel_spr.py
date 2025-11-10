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
        """
        Функция для парсинга Excel-файла и сохранения данных в JSON
        :param output_dir:
        :param output_file:
        :return:
        """
        if not self.all_data:  # Парсим только если данные еще не загружены
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

    def format_tel(self, tel):
        """Форматирование телефонного номера"""
        if not isinstance(tel, (str, int)):
            return tel

        tel_str = str(tel).strip()

        # Убираем все нецифровые символы
        digits = ''.join(filter(str.isdigit, tel_str))

        if len(digits) == 6:
            return f"{digits[0:2]}-{digits[2:4]}-{digits[4:]}"
        elif len(digits) == 5:
            return f"{digits[0:1]}-{digits[1:3]}-{digits[3:]}"
        else:
            return tel_str  # Возвращаем как есть, если не подходит под форматы


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

                    # # ---------------------------------------------------------------------------------------------------------
                    num_tel = row.get("Unnamed: 0", '')
                    num_tel = str(num_tel).strip()
                    if "-" not in num_tel and len(num_tel) > 0 and num_tel.isdigit():
                        num_tel = self.format_tel(num_tel)
                        row["Unnamed: 0"] = num_tel
                    # # ---------------------------------------------------------------------------------------------------------

                    if row.get("Unnamed: 8", '') != '':
                        # Создаем новую строку с замененными ключами
                        new_row = {}
                        for old_key, value in row.items():
                            # Если ключ есть в target_map, используем новое название, иначе оставляем старый
                            new_key = self.target_map.get(old_key, old_key)
                            new_row[new_key] = str(value).strip()
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

                    # ---------------------------------------------------------------------------------------------------------
                    num_tel = row.get("Unnamed: 1", '')
                    num_tel = str(num_tel).strip()
                    if "-" not in num_tel and len(num_tel) > 0 and num_tel.isdigit():
                        num_tel = self.format_tel(num_tel)
                        row["Unnamed: 1"] = num_tel
                    # ---------------------------------------------------------------------------------------------------------

                    if row.get("Unnamed: 6", '') != '':
                        # Создаем новую строку с замененными ключами
                        new_row = {}
                        new_row['отдел'] = otdel
                        for old_key, value in row.items():
                            new_key = self.target_map.get(old_key, old_key)
                            new_row[new_key] = str(value).strip()

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


def load_xls(path_xls):
    try:
        # Создаем объект один раз для каждого типа
        ks = SprKs(path_xls)
        ks_data = ks.obrabotka_ks()  # Получаем данные
        ks.save_to_json()  # Сохраняем

        osfr = SprOsfr(path_xls)
        osfr_data = osfr.obrabotka_osfr()
        osfr.save_to_json(output_file='osfr.json')

        print("Обработка завершена успешно!")
        return ks_data, osfr_data  # Возвращаем данные для дальнейшего использования

    except FileNotFoundError as e:
        print(f"Ошибка: {e}")
        return None, None
    except Exception as e:
        print(f"Неожиданная ошибка: {e}")
        return None, None