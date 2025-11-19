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
    
    def extract_technical_data(self, pdf_path: str, page_number: int = 2) -> Dict[str, Tuple[str, str]]:
        """
        Извлечение таблицы "Основные технические данные" из паспорта
        
        Args:
            pdf_path: путь к PDF файлу
            page_number: номер страницы с таблицей (по умолчанию 2)
            
        Returns:
            Словарь {параметр: (значение, единица_измерения)}
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF файл не найден: {pdf_path}")
        
        technical_data = {}
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Проверяем, есть ли нужная страница
                if len(pdf.pages) < page_number:
                    raise Exception(f"В PDF файле меньше {page_number} страниц")
                
                # Берём нужную страницу (индексация с 0)
                page = pdf.pages[page_number - 1]
                
                # Извлекаем таблицы
                tables = page.extract_tables()
                
                if not tables:
                    # Пробуем извлечь текст и найти таблицу вручную
                    text = page.extract_text()
                    technical_data = self._parse_text_table(text)
                else:
                    # Обрабатываем таблицы
                    for table in tables:
                        parsed_data = self._parse_table(table)
                        if parsed_data:
                            technical_data.update(parsed_data)
                
        except Exception as e:
            raise Exception(f"Ошибка при парсинге PDF: {str(e)}")
        
        # Удаляем параметр "размер зоны приземления" согласно ТЗ
        keys_to_remove = []
        for key in technical_data.keys():
            if 'зон' in key.lower() and 'приземлен' in key.lower():
                keys_to_remove.append(key)
        
        for key in keys_to_remove:
            del technical_data[key]
        
        return technical_data
    
    @staticmethod
    def _parse_table(table: List[List[str]]) -> Dict[str, Tuple[str, str]]:
        """
        Парсинг таблицы, извлечённой из PDF
        
        Args:
            table: таблица в виде списка списков
            
        Returns:
            Словарь {параметр: (значение, единица)}
        """
        data = {}
        
        for row in table:
            if not row or len(row) < 2:
                continue
            
            # Очищаем ячейки от None и пробелов
            cleaned_row = [str(cell).strip() if cell else '' for cell in row]
            
            # Пропускаем заголовки и пустые строки
            if not cleaned_row[0] or cleaned_row[0].lower() in ['параметр', 'наименование', '№']:
                continue
            
            # Первый столбец - название параметра
            param_name = cleaned_row[0]
            
            # Второй столбец - значение
            value = cleaned_row[1] if len(cleaned_row) > 1 else ''
            
            # Третий столбец - единица измерения (если есть)
            unit = cleaned_row[2] if len(cleaned_row) > 2 else ''
            
            # Сохраняем данные
            if param_name and value:
                data[param_name] = (value, unit)
        
        return data
    
    @staticmethod
    def _parse_text_table(text: str) -> Dict[str, Tuple[str, str]]:
        """
        Парсинг таблицы из текста (если extract_tables не сработал)
        
        Args:
            text: текст страницы
            
        Returns:
            Словарь {параметр: (значение, единица)}
        """
        data = {}
        
        # Ищем строки, которые выглядят как записи таблицы
        lines = text.split('\n')
        
        for line in lines:
            # Пропускаем короткие строки и заголовки
            if len(line.strip()) < 5:
                continue
            
            # Простой парсинг: ищем строки с числами
            # Это упрощённая версия, может требовать доработки
            parts = line.split()
            if len(parts) >= 2:
                # Берём первую часть как параметр, последние - как значения
                param = ' '.join(parts[:-2]) if len(parts) > 2 else parts[0]
                value = parts[-2] if len(parts) > 1 else parts[-1]
                unit = parts[-1] if len(parts) > 2 else ''
                
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
