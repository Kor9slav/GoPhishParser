import csv
import json
import argparse
from collections import defaultdict
from openpyxl import Workbook
import os
import textwrap

def print_banner():
    banner = r"""
 _____      ______ _     _     _    ______                   
|  __ \     | ___ \ |   (_)   | |   | ___ \                  
| |  \/ ___ | |_/ / |__  _ ___| |__ | |_/ /_ _ _ __ ___  ___ 
| | __ / _ \|  __/| '_ \| / __| '_ \|  __/ _` | '__/ __|/ _ \
| |_\ \ (_) | |   | | | | \__ \ | | | | | (_| | |  \__ \  __/
 \____/\___/\_|   |_| |_|_|___/_| |_\_|  \__,_|_|  |___/\___|
 					   GoPhish CSV Parser
    """
    print(banner)

def mask_password(pw):
    if not pw:
        return ''
    if len(pw) <= 2:
        return pw[0] + '*' * (len(pw) - 1)
    if len(pw) <= 4:
        return pw[0] + '*' * (len(pw) - 2) + pw[-1]
    return pw[:2] + '*' * (len(pw) - 4) + pw[-2:]

def write_excel(data, path, masked=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "Gophish Results"
    ws.append(["Email", "Password", "Переход по ссылке"])

    for email, info in sorted(data.items()):
        password = mask_password(info['password']) if masked else info['password']
        ws.append([email, password, '+' if info['clicked'] else ''])

    wb.save(path)
    print(f"[+] Сохранён файл: {path}")

def parse_gophish_csv(input_csv_path, output_xlsx_path, generate_masked=False):
    user_data = defaultdict(lambda: {'clicked': False, 'password': ''})

    with open(input_csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            email = row['email'].strip()
            if not email:
                continue

            message = row['message'].strip()
            details = row['details'].strip()

            if message == 'Clicked Link':
                user_data[email]['clicked'] = True
            elif message == 'Submitted Data':
                try:
                    details_json = json.loads(details)
                    password_list = details_json.get('payload', {}).get('password', [])
                    if password_list:
                        user_data[email]['password'] = password_list[0]
                except json.JSONDecodeError:
                    pass

    # Основной файл
    write_excel(user_data, output_xlsx_path, masked=False)

    # Маскированный файл
    if generate_masked:
        base, ext = os.path.splitext(output_xlsx_path)
        masked_path = f"{base}_masked{ext}"
        write_excel(user_data, masked_path, masked=True)

def main():
    print_banner()

    description = textwrap.dedent("""
        GoParse — инструмент для парсинга CSV-файлов с результатами Gophish.
        Извлекает информацию о кликах и введённых паролях, экспортирует в Excel.

        Примеры использования:
          python gophish_parse.py -t results.csv -o parsed.xlsx
          python gophish_parse.py -t results.csv -o parsed.xlsx --hide
    """)

    parser = argparse.ArgumentParser(
        description=description,
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('-t', '--target', required=True, help='Путь к входному CSV-файлу от Gophish')
    parser.add_argument('-o', '--output', required=True, help='Путь к выходному XLSX-файлу')
    parser.add_argument('--hide', action='store_true', help='Создать также второй файл с маскированными паролями')

    args = parser.parse_args()
    parse_gophish_csv(args.target, args.output, generate_masked=args.hide)

if __name__ == '__main__':
    main()
