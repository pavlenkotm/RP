"""
Модуль для извлечения данных из PDF паспортов изделий
"""
import os
import glob
from typing import Dict, List, Optional, Tuple
import pdfplumber
import re


class PDFParser:
    """Класс для парсинга PDF паспортов"""
    
    def __init__(self, passports_dir: str, pattern: str = "*{ART}*.pdf"):
        """
        Инициализация парсера PDF
        
        Args:
            passports_dir: директория с паспортами
            pattern: шаблон поиска файла паспорта
        """
        self.passports_dir = passports_dir
        self.pattern = pattern
    
    def find_passport(self, article: str) -> Optional[str]:
        """
        Поиск файла паспорта по артикулу
        
        Args:
            article: артикул изделия
            
        Returns:
            Путь к файлу паспорта или None
        """
        # Формируем паттерн поиска
        search_pattern = self.pattern.replace('{ART}', article)
        search_path = os.path.join(self.passports_dir, search_pattern)
        
        # Ищем файлы
        files = glob.glob(search_path)
        
        if files:
            return files[0]  # Возвращаем первый найденный файл
        
        return None
    
    def extract_technical_data(self, pdf_path: str) -> Dict[str, Tuple[str, str]]:
        """
        Извлечение таблицы "Основные технические данные" из паспорта
        
        Args:
            pdf_path: путь к PDF файлу
            
        Returns:
            Словарь {параметр: (значение, единица_измерения)}
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF файл не найден: {pdf_path}")
        
        technical_data = {}
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Ищем страницу с "Основные технические данные"
                target_page = None
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    if "Основные технические данные" in text or "ОСНОВНЫЕ ТЕХНИЧЕСКИЕ ДАННЫЕ" in text:
                        target_page = page
                        break
                
                if not target_page:
                    # Если не нашли по заголовку, пробуем страницу 4 (индекс 3)
                    if len(pdf.pages) >= 4:
                        target_page = pdf.pages[3]
                
                if target_page:
                    # Извлекаем таблицы
                    tables = target_page.extract_tables()
                    
                    # Ищем таблицу с техническими данными
                    for table in tables:
                        if len(table) > 1:
                            # Проверяем заголовок таблицы
                            header = table[0] if table[0] else []
                            header_text = ' '.join([str(cell) for cell in header if cell]).upper()
                            
                            if 'ПАРАМЕТР' in header_text or 'ЗНАЧЕНИЕ' in header_text:
                                # Это наша таблица
                                technical_data = self._parse_technical_table(table)
                                break
                    
                    # Если не нашли таблицу, пробуем извлечь из текста
                    if not technical_data:
                        text = target_page.extract_text()
                        technical_data = self._parse_text_data(text)
                
        except Exception as e:
            raise Exception(f"Ошибка при парсинге PDF: {str(e)}")
        
        # Удаляем параметр "размер зоны приземления" согласно ТЗ
        keys_to_remove = []
        for key in technical_data.keys():
            key_lower = key.lower()
            if ('зон' in key_lower and 'приземлен' in key_lower) or 'зоны приземления' in key_lower:
                keys_to_remove.append(key)
        
        for key in keys_to_remove:
            del technical_data[key]
        
        return technical_data
    
    @staticmethod
    def _parse_technical_table(table: List[List[str]]) -> Dict[str, Tuple[str, str]]:
        """
        Парсинг таблицы технических данных
        
        Args:
            table: таблица в виде списка списков
            
        Returns:
            Словарь {параметр: (значение, единица)}
        """
        data = {}
        
        # Пропускаем заголовок (первая строка)
        for row in table[1:]:
            if not row or len(row) < 2:
                continue
            
            # Очищаем ячейки
            param_cell = str(row[0]) if row[0] else ''
            value_cell = str(row[1]) if row[1] else ''
            
            # Пропускаем пустые строки
            if not param_cell.strip() or not value_cell.strip():
                continue
            
            # Извлекаем параметр
            param_name = param_cell.strip()
            # Убираем из названия единицы измерения (они в скобках)
            param_name = re.sub(r'\([^)]+\)', '', param_name).strip()
            
            # Извлекаем значение
            value = value_cell.strip()
            
            # Пытаемся определить единицу измерения
            unit = ''
            if 'мм' in param_cell.lower():
                unit = 'мм'
            elif 'кг' in param_cell.lower():
                unit = 'кг'
            elif 'м' in param_cell.lower():
                unit = 'м'
            
            # Сохраняем
            if param_name and value:
                data[param_name] = (value, unit)
        
        return data
    
    @staticmethod
    def _parse_text_data(text: str) -> Dict[str, Tuple[str, str]]:
        """
        Парсинг технических данных из текста (резервный метод)
        
        Args:
            text: текст страницы
            
        Returns:
            Словарь {параметр: (значение, единица)}
        """
        data = {}
        
        # Ищем строки вида "Параметр, единица: значение" или "Параметр значение"
        lines = text.split('\n')
        
        for line in lines:
            # Ищем строки с размерами и параметрами
            if any(keyword in line.lower() for keyword in ['длина', 'ширина', 'высота', 'масса']):
                # Пытаемся извлечь параметр и значение
                parts = re.split(r'[:\s]{2,}', line)
                if len(parts) >= 2:
                    param = parts[0].strip()
                    value = parts[1].strip()
                    
                    # Определяем единицу
                    unit = ''
                    if 'мм' in line:
                        unit = 'мм'
                    elif 'кг' in line:
                        unit = 'кг'
                    elif 'м' in line:
                        unit = 'м'
                    
                    # Убираем единицу из значения
                    value = re.sub(r'[а-яА-Я]+', '', value).strip()
                    
                    # Убираем скобки из параметра
                    param = re.sub(r'\([^)]+\)', '', param).strip()
                    
                    if param and value:
                        data[param] = (value, unit)
        
        return data
    
    def extract_all_data(self, article: str) -> Optional[Dict[str, Tuple[str, str]]]:
        """
        Полный процесс: найти паспорт и извлечь данные
        
        Args:
            article: артикул изделия
            
        Returns:
            Словарь с техническими данными или None при ошибке
        """
        # Ищем паспорт
        passport_path = self.find_passport(article)
        
        if not passport_path:
            return None
        
        # Извлекаем данные
        try:
            return self.extract_technical_data(passport_path)
        except Exception as e:
            raise Exception(f"Ошибка извлечения данных из {passport_path}: {str(e)}")
