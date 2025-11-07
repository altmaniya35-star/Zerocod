#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Скрипт для генерации PDF-счетов из CSV и JSON файлов
Использует HTML-шаблоны и библиотеку WeasyPrint для создания PDF
"""

import os
import json
import csv
import sys
import platform
import subprocess
from pathlib import Path
from typing import Dict, List, Any, Optional

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from weasyprint import HTML, CSS
    from weasyprint.text.fonts import FontConfiguration
except ImportError:
    print("Ошибка: библиотека WeasyPrint не установлена.")
    print("Установите её командой: pip install weasyprint")
    sys.exit(1)


class InvoiceGenerator:
    """Класс для генерации PDF-счетов из данных"""
    
    def __init__(self, data_dir: str = "data", templates_dir: str = "templates", output_dir: str = "output"):
        """
        Инициализация генератора счетов
        
        Args:
            data_dir: Директория с файлами данных (CSV, JSON)
            templates_dir: Директория с HTML-шаблонами
            output_dir: Директория для сохранения PDF-файлов
        """
        self.data_dir = Path(data_dir)
        self.templates_dir = Path(templates_dir)
        self.output_dir = Path(output_dir)
        
        # Создаём выходную директорию, если её нет
        self.output_dir.mkdir(exist_ok=True)
        
        # Настройка шрифта для поддержки кириллицы
        self.font_config = FontConfiguration()
    
    def get_data_files(self) -> List[Path]:
        """
        Получает список всех CSV и JSON файлов в директории данных
        
        Returns:
            Список путей к файлам данных
        """
        data_files = []
        
        if not self.data_dir.exists():
            return []
        
        # Ищем CSV файлы
        csv_files = list(self.data_dir.glob("*.csv"))
        # Ищем JSON файлы
        json_files = list(self.data_dir.glob("*.json"))
        
        data_files = csv_files + json_files
        return sorted(data_files)
    
    def get_template_files(self) -> List[Path]:
        """
        Получает список всех HTML-шаблонов в директории шаблонов
        
        Returns:
            Список путей к HTML-шаблонам
        """
        if not self.templates_dir.exists():
            return []
        
        template_files = list(self.templates_dir.glob("*.html"))
        return sorted(template_files)
    
    def load_csv_data(self, file_path: Path) -> List[Dict[str, Any]]:
        """
        Загружает данные из CSV файла
        
        Args:
            file_path: Путь к CSV файлу
            
        Returns:
            Список словарей с данными
        """
        if pd is not None:
            # Используем pandas, если доступен
            df = pd.read_csv(file_path, encoding='utf-8')
            return df.to_dict('records')
        else:
            # Используем стандартную библиотеку csv
            data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    data.append(row)
            return data
    
    def load_json_data(self, file_path: Path) -> List[Dict[str, Any]]:
        """
        Загружает данные из JSON файла
        
        Args:
            file_path: Путь к JSON файлу
            
        Returns:
            Список словарей с данными
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Если это не список, оборачиваем в список
        if not isinstance(data, list):
            data = [data]
        
        return data
    
    def load_data_file(self, file_path: Path) -> List[Dict[str, Any]]:
        """
        Загружает данные из файла (CSV или JSON)
        
        Args:
            file_path: Путь к файлу данных
            
        Returns:
            Список словарей с данными
        """
        if file_path.suffix.lower() == '.csv':
            return self.load_csv_data(file_path)
        elif file_path.suffix.lower() == '.json':
            return self.load_json_data(file_path)
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {file_path.suffix}")
    
    def get_invoice_ids(self, data: List[Dict[str, Any]], file_path: Path) -> List[Any]:
        """
        Извлекает список уникальных invoice_id из данных
        
        Args:
            data: Список словарей с данными
            file_path: Путь к файлу (для определения структуры)
            
        Returns:
            Список уникальных invoice_id
        """
        invoice_ids = set()
        
        # Для CSV файлов invoice может быть в каждой строке
        if file_path.suffix.lower() == '.csv':
            for row in data:
                if 'invoice_id' in row:
                    invoice_ids.add(row['invoice_id'])
        # Для JSON файлов invoice может быть в корне или в items
        elif file_path.suffix.lower() == '.json':
            for item in data:
                if 'invoice_id' in item:
                    invoice_ids.add(item['invoice_id'])
        
        return sorted(list(invoice_ids))
    
    def get_invoice_data(self, data: List[Dict[str, Any]], invoice_id: Any, file_path: Path) -> Dict[str, Any]:
        """
        Получает данные конкретного счёта
        
        Args:
            data: Все данные из файла
            invoice_id: ID счёта
            file_path: Путь к файлу данных
            
        Returns:
            Словарь с данными счёта
        """
        if file_path.suffix.lower() == '.csv':
            # Для CSV собираем все строки с данным invoice_id
            invoice_items = []
            invoice_date = None
            customer_id = None
            
            for row in data:
                if str(row.get('invoice_id')) == str(invoice_id):
                    invoice_items.append(row)
                    if invoice_date is None:
                        invoice_date = row.get('date')
                    if customer_id is None:
                        customer_id = row.get('customer_id')
            
            return {
                'invoice_id': invoice_id,
                'customer_id': customer_id,
                'date': invoice_date,
                'items': invoice_items
            }
        else:  # JSON
            for item in data:
                if str(item.get('invoice_id')) == str(invoice_id):
                    return item
        
        raise ValueError(f"Счёт с ID {invoice_id} не найден")
    
    def load_customer_data(self, customer_id: Any) -> Optional[Dict[str, Any]]:
        """
        Загружает данные покупателя по ID
        
        Args:
            customer_id: ID покупателя
            
        Returns:
            Словарь с данными покупателя или None
        """
        # Пробуем загрузить из CSV
        csv_file = self.data_dir / "customer.csv"
        if csv_file.exists():
            customers = self.load_csv_data(csv_file)
            for customer in customers:
                if str(customer.get('customer_id')) == str(customer_id):
                    return customer
        
        # Пробуем загрузить из JSON
        json_file = self.data_dir / "customer.json"
        if json_file.exists():
            customers = self.load_json_data(json_file)
            for customer in customers:
                if str(customer.get('customer_id')) == str(customer_id):
                    return customer
        
        return None
    
    def load_product_data(self, product_id: Any) -> Optional[Dict[str, Any]]:
        """
        Загружает данные товара по ID
        
        Args:
            product_id: ID товара
            
        Returns:
            Словарь с данными товара или None
        """
        # Пробуем загрузить из CSV
        csv_file = self.data_dir / "product.csv"
        if csv_file.exists():
            products = self.load_csv_data(csv_file)
            for product in products:
                if str(product.get('product_id')) == str(product_id):
                    return product
        
        # Пробуем загрузить из JSON
        json_file = self.data_dir / "product.json"
        if json_file.exists():
            products = self.load_json_data(json_file)
            for product in products:
                if str(product.get('product_id')) == str(product_id):
                    return product
        
        return None
    
    def prepare_invoice_items(self, invoice_data: Dict[str, Any], file_path: Path) -> List[Dict[str, Any]]:
        """
        Подготавливает список товаров для счёта с полной информацией
        
        Args:
            invoice_data: Данные счёта
            file_path: Путь к файлу данных
            
        Returns:
            Список словарей с полной информацией о товарах
        """
        items = []
        
        if file_path.suffix.lower() == '.csv':
            # Для CSV items уже в invoice_data['items']
            for item in invoice_data['items']:
                product_id = item.get('product_id')
                quantity = int(item.get('quantity', 1))
                
                product = self.load_product_data(product_id)
                if product:
                    price = float(product.get('price', 0))
                    total = price * quantity
                    
                    items.append({
                        'product_id': product_id,
                        'name': product.get('name', 'Неизвестный товар'),
                        'quantity': quantity,
                        'price': price,
                        'total': total
                    })
        else:  # JSON
            # Для JSON items могут быть в invoice_data['items']
            invoice_items = invoice_data.get('items', [])
            for item in invoice_items:
                product_id = item.get('product_id')
                quantity = int(item.get('quantity', 1))
                
                product = self.load_product_data(product_id)
                if product:
                    price = float(product.get('price', 0))
                    total = price * quantity
                    
                    items.append({
                        'product_id': product_id,
                        'name': product.get('name', 'Неизвестный товар'),
                        'quantity': quantity,
                        'price': price,
                        'total': total
                    })
        
        return items
    
    def generate_html(self, template_path: Path, invoice_data: Dict[str, Any], 
                     customer_data: Dict[str, Any], items: List[Dict[str, Any]]) -> str:
        """
        Генерирует HTML из шаблона с подстановкой данных
        Разбивает товары на группы по 10 записей на страницу
        
        Args:
            template_path: Путь к HTML-шаблону
            invoice_data: Данные счёта
            customer_data: Данные покупателя
            items: Список товаров
            
        Returns:
            HTML-строка с подставленными данными
        """
        # Читаем шаблон
        with open(template_path, 'r', encoding='utf-8') as f:
            html_template = f.read()
        
        # Подставляем данные счёта
        html = html_template.replace('{{invoice_id}}', str(invoice_data.get('invoice_id', '')))
        html = html.replace('{{invoice_date}}', str(invoice_data.get('date', '')))
        
        # Подставляем данные покупателя
        html = html.replace('{{customer_name}}', customer_data.get('name', ''))
        html = html.replace('{{customer_email}}', customer_data.get('email', ''))
        html = html.replace('{{customer_phone}}', customer_data.get('phone', ''))
        html = html.replace('{{customer_address}}', customer_data.get('address', ''))
        
        # Вычисляем итоговую сумму
        total_amount = sum(item['total'] for item in items)
        
        # Разбиваем товары на группы по 10 записей
        items_per_page = 10
        tables_html = ""
        
        for page_num in range(0, len(items), items_per_page):
            page_items = items[page_num:page_num + items_per_page]
            is_first_page = (page_num == 0)
            
            # Для первой страницы - обычная таблица
            # Для последующих - добавляем разрыв страницы и повторяем шапку
            if not is_first_page:
                # Добавляем разрыв страницы и повторяем шапку счёта
                tables_html += f"""
    <div class="page-break"></div>
    <div class="header-repeat">
        <h2>Счёт №{invoice_data.get('invoice_id', '')}</h2>
        <div class="header-repeat-info">
            <div class="header-repeat-info-row">
                <div class="header-repeat-info-cell header-repeat-info-label">Дата:</div>
                <div class="header-repeat-info-cell">{invoice_data.get('date', '')}</div>
            </div>
        </div>
        <div class="header-repeat-customer">
            <h3>Покупатель:</h3>
            <div class="header-repeat-customer-details">
                <p><strong>{customer_data.get('name', '')}</strong></p>
                <p>Email: {customer_data.get('email', '')}</p>
                <p>Телефон: {customer_data.get('phone', '')}</p>
                <p>Адрес: {customer_data.get('address', '')}</p>
            </div>
        </div>
    </div>
"""
            
            # Генерируем таблицу для текущей страницы
            items_rows = ""
            for idx, item in enumerate(page_items, 1):
                global_idx = page_num + idx
                items_rows += f"""
            <tr>
                <td class="text-center">{global_idx}</td>
                <td>{item['name']}</td>
                <td class="text-center">{item['quantity']}</td>
                <td class="text-right">{item['price']:,.2f} ₽</td>
                <td class="text-right">{item['total']:,.2f} ₽</td>
            </tr>
"""
            
            # Добавляем таблицу
            tables_html += f"""
    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th style="width: 5%;">№</th>
                    <th style="width: 45%;">Наименование товара</th>
                    <th style="width: 10%;" class="text-center">Кол-во</th>
                    <th style="width: 15%;" class="text-right">Цена за шт.</th>
                    <th style="width: 15%;" class="text-right">Сумма</th>
                </tr>
            </thead>
            <tbody>
                {items_rows}
            </tbody>
        </table>
    </div>
"""
        
        # Подставляем таблицы и итоговую сумму
        html = html.replace('{{tables}}', tables_html)
        html = html.replace('{{total_amount}}', f"{total_amount:,.2f}")
        
        return html
    
    def generate_pdf(self, html_content: str, output_path: Path):
        """
        Генерирует PDF из HTML-контента
        
        Args:
            html_content: HTML-контент
            output_path: Путь для сохранения PDF
        """
        # CSS для поддержки кириллицы и пагинации
        # Разрывы страниц контролируются через класс .page-break
        css = CSS(string='''
            @page {
                size: A4;
                margin: 2cm 2cm 3cm 2cm;
            }
            
            body {
                font-family: 'DejaVu Sans', Arial, sans-serif;
            }
            
            .page-break {
                page-break-before: always;
            }
            
            table {
                page-break-inside: avoid;
            }
            
            thead {
                display: table-header-group;
            }
            
            tbody tr {
                page-break-inside: avoid;
            }
            
            .total-section {
                page-break-inside: avoid;
            }
        ''', font_config=self.font_config)
        
        # Генерируем PDF
        HTML(string=html_content).write_pdf(
            output_path,
            stylesheets=[css],
            font_config=self.font_config
        )
    
    def open_pdf(self, pdf_path: Path):
        """
        Открывает PDF файл в системной программе
        
        Args:
            pdf_path: Путь к PDF файлу
        """
        system = platform.system()
        
        if system == 'Windows':
            os.startfile(str(pdf_path))
        elif system == 'Darwin':  # macOS
            subprocess.run(['open', str(pdf_path)])
        else:  # Linux
            subprocess.run(['xdg-open', str(pdf_path)])
    
    def run(self):
        """Основной метод для запуска генератора"""
        print("=" * 60)
        print("ГЕНЕРАТОР PDF-СЧЕТОВ")
        print("=" * 60)
        print()
        
        # Получаем список файлов данных
        data_files = self.get_data_files()
        if not data_files:
            print("Ошибка: не найдено файлов данных в директории", self.data_dir)
            return
        
        # Получаем список шаблонов
        template_files = self.get_template_files()
        if not template_files:
            print("Ошибка: не найдено HTML-шаблонов в директории", self.templates_dir)
            return
        
        # Выводим список файлов данных
        print("Доступные файлы с данными:")
        print("-" * 60)
        for idx, file_path in enumerate(data_files, 1):
            print(f"{idx}. {file_path.name}")
        print()
        
        # Выводим список шаблонов
        print("Доступные HTML-шаблоны:")
        print("-" * 60)
        for idx, template_path in enumerate(template_files, 1):
            print(f"{idx}. {template_path.name}")
        print()
        
        # Выбор файла данных
        while True:
            try:
                data_choice = input(f"Выберите файл данных (1-{len(data_files)}): ").strip()
                data_idx = int(data_choice) - 1
                if 0 <= data_idx < len(data_files):
                    selected_data_file = data_files[data_idx]
                    break
                else:
                    print("Неверный номер. Попробуйте снова.")
            except ValueError:
                print("Введите число.")
            except KeyboardInterrupt:
                print("\nПрервано пользователем.")
                return
        
        # Выбор шаблона
        while True:
            try:
                template_choice = input(f"Выберите HTML-шаблон (1-{len(template_files)}): ").strip()
                template_idx = int(template_choice) - 1
                if 0 <= template_idx < len(template_files):
                    selected_template = template_files[template_idx]
                    break
                else:
                    print("Неверный номер. Попробуйте снова.")
            except ValueError:
                print("Введите число.")
            except KeyboardInterrupt:
                print("\nПрервано пользователем.")
                return
        
        print()
        print("Загрузка данных...")
        
        # Загружаем данные
        try:
            data = self.load_data_file(selected_data_file)
        except Exception as e:
            print(f"Ошибка при загрузке данных: {e}")
            return
        
        # Получаем список invoice_id
        invoice_ids = self.get_invoice_ids(data, selected_data_file)
        if not invoice_ids:
            print("Ошибка: не найдено счетов в файле данных.")
            return
        
        # Выводим список счетов
        print()
        print("Доступные счета (по invoice_id):")
        print("-" * 60)
        for idx, invoice_id in enumerate(invoice_ids, 1):
            print(f"{idx}. Счёт №{invoice_id}")
        print()
        
        # Выбор счёта
        while True:
            try:
                invoice_choice = input(f"Выберите счёт (1-{len(invoice_ids)}): ").strip()
                invoice_idx = int(invoice_choice) - 1
                if 0 <= invoice_idx < len(invoice_ids):
                    selected_invoice_id = invoice_ids[invoice_idx]
                    break
                else:
                    print("Неверный номер. Попробуйте снова.")
            except ValueError:
                print("Введите число.")
            except KeyboardInterrupt:
                print("\nПрервано пользователем.")
                return
        
        print()
        print("Генерация PDF...")
        
        # Получаем данные счёта
        try:
            invoice_data = self.get_invoice_data(data, selected_invoice_id, selected_data_file)
        except Exception as e:
            print(f"Ошибка при получении данных счёта: {e}")
            return
        
        # Получаем данные покупателя
        customer_id = invoice_data.get('customer_id')
        if not customer_id:
            print("Ошибка: не указан customer_id в данных счёта.")
            return
        
        customer_data = self.load_customer_data(customer_id)
        if not customer_data:
            print(f"Ошибка: покупатель с ID {customer_id} не найден.")
            return
        
        # Подготавливаем товары
        items = self.prepare_invoice_items(invoice_data, selected_data_file)
        if not items:
            print("Ошибка: не найдено товаров в счёте.")
            return
        
        # Генерируем HTML
        try:
            html_content = self.generate_html(selected_template, invoice_data, customer_data, items)
        except Exception as e:
            print(f"Ошибка при генерации HTML: {e}")
            return
        
        # Генерируем PDF
        output_filename = f"invoice_{selected_invoice_id}.pdf"
        output_path = self.output_dir / output_filename
        
        try:
            self.generate_pdf(html_content, output_path)
            print(f"PDF успешно создан: {output_path}")
        except Exception as e:
            print(f"Ошибка при генерации PDF: {e}")
            return
        
        # Открываем PDF
        try:
            print("Открытие PDF...")
            self.open_pdf(output_path)
        except Exception as e:
            print(f"Не удалось открыть PDF автоматически: {e}")
            print(f"PDF сохранён в: {output_path}")


def main():
    """Главная функция"""
    generator = InvoiceGenerator()
    generator.run()


if __name__ == "__main__":
    main()
