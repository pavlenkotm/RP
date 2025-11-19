"""
Модуль для извлечения данных из PDF паспортов - ФИНАЛЬНАЯ ВЕРСИЯ
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
        """
        self.passports_dir = passports_dir
        self.pattern = pattern
    
    def find_passport(self, article: str) -> Optional[str]:
        """
        Поиск файла паспорта по артикулу
        """
        search_pattern = self.pattern.replace('{ART}', article)
        search_path = os.path.join(self.passports_dir, search_pattern)
        
        files = glob.glob(search_path)
        
        if files:
            return files[0]
        
        return None
    
    def extract_technical_data(self, pdf_path: str) -> Dict[str, Tuple[str, str]]:
        """
        Извлечение таблицы "Основные технические данные"
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF файл не найден: {pdf_path}")
        
        technical_data = {}
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Стратегия 1: Ищем страницу с таблицей технических данных
                for page_num, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    
                    # Ищем таблицу с техническими данными
                    for table in tables:
                        if len(table) > 1:
                            # Проверяем заголовок таблицы
                            header_row = table[0] if table[0] else []
                            header_text = ' '.join([str(cell).upper() for cell in header_row if cell])
                            
                            # Это таблица с техническими данными?
                            if 'ПАРАМЕТР' in header_text and 'ЗНАЧЕНИЕ' in header_text:
                                technical_data = self._parse_technical_table(table)
                                if technical_data:
                                    break
                    
                    if technical_data:
                        break
                
                # Стратегия 2: Если не нашли по таблице, ищем страницу с текстом и парсим
                if not technical_data:
                    # Пробуем страницы 3-5 (обычно там технические данные)
                    for page_idx in [3, 4, 2]:
                        if page_idx < len(pdf.pages):
                            page = pdf.pages[page_idx]
                            text = page.extract_text() or ""
                            
                            # Проверяем что это действительно страница с данными
                            if '2. Основные технические данные' in text or 'Основные технические данные' in text:
                                # Ищем таблицу на этой странице
                                tables = page.extract_tables()
                                for table in tables:
                                    if len(table) > 2:
                                        result = self._parse_technical_table(table)
                                        if result:
                                            technical_data = result
                                            break
                                
                                # Если таблицу не нашли, парсим из текста
                                if not technical_data:
                                    technical_data = self._parse_text_data(text)
                                
                                if technical_data:
                                    break
                
        except Exception as e:
            raise Exception(f"Ошибка при парсинге PDF: {str(e)}")
        
        # Удаляем параметр "размер зоны приземления"
        keys_to_remove = []
        for key in technical_data.keys():
            key_lower = key.lower()
            if 'зон' in key_lower and 'приземлен' in key_lower:
                keys_to_remove.append(key)
        
        for key in keys_to_remove:
            del technical_data[key]
        
        return technical_data
    
    @staticmethod
    def _parse_technical_table(table: List[List[str]]) -> Dict[str, Tuple[str, str]]:
        """
        Парсинг таблицы технических данных
        """
        data = {}
        
        # Определяем, есть ли заголовок
        has_header = False
        if table[0]:
            header_text = ' '.join([str(c).upper() for c in table[0] if c])
            if 'ПАРАМЕТР' in header_text or 'НАИМЕНОВАНИЕ' in header_text:
                has_header = True
        
        # Начинаем со второй строки если есть заголовок
        start_row = 1 if has_header else 0
        
        for row in table[start_row:]:
            if not row or len(row) < 2:
                continue
            
            # Извлекаем параметр и значение
            param_cell = str(row[0]) if row[0] else ''
            value_cell = str(row[1]) if row[1] else ''
            
            # Пропускаем пустые и заголовочные строки
            param_clean = param_cell.strip()
            value_clean = value_cell.strip()
            
            if not param_clean or not value_clean:
                continue
            
            if 'ПАРАМЕТР' in param_clean.upper() or 'ЗНАЧЕНИЕ' in value_clean.upper():
                continue
            
            # Убираем единицы измерения и скобки из названия параметра
            param_name = re.sub(r'\([^)]+\)', '', param_clean).strip()
            param_name = re.sub(r'\n.*', '', param_name).strip()  # Убираем переносы строк
            
            # Определяем единицу измерения
            unit = ''
            param_lower = param_cell.lower()
            if 'мм' in param_lower:
                unit = 'мм'
            elif 'кг' in param_lower:
                unit = 'кг'
            elif 'м' in param_lower and 'мм' not in param_lower:
                unit = 'м'
            
            # Очищаем значение
            value = value_clean.replace('\n', ' ').strip()
            
            # Сохраняем
            if param_name and value and len(param_name) > 2:
                data[param_name] = (value, unit)
        
        return data
    
    @staticmethod
    def _parse_text_data(text: str) -> Dict[str, Tuple[str, str]]:
        """
        Парсинг данных из текста (резервный метод)
        """
        data = {}
        
        # Ищем строки с параметрами
        lines = text.split('\n')
        
        for line in lines:
            # Пропускаем короткие строки
            if len(line.strip()) < 10:
                continue
            
            # Ищем строки с размерами
            if any(keyword in line.lower() for keyword in ['длина', 'ширина', 'высота', 'масса']):
                # Пытаемся извлечь параметр и значение
                # Формат: "Длина, мм 10 222" или "Длина, мм\n10 222"
                parts = re.split(r'[\t\s]{2,}', line)
                
                if len(parts) >= 2:
                    param = parts[0].strip()
                    value = parts[-1].strip()
                    
                    # Определяем единицу
                    unit = ''
                    if 'мм' in param.lower():
                        unit = 'мм'
                    elif 'кг' in param.lower():
                        unit = 'кг'
                    
                    # Убираем скобки и единицы из параметра
                    param = re.sub(r'\([^)]+\)', '', param)
                    param = re.sub(r',\s*(мм|кг|м)\s*', '', param).strip()
                    
                    # Убираем буквы из значения
                    value = re.sub(r'[а-яА-Я]+', '', value).strip()
                    
                    if param and value:
                        data[param] = (value, unit)
        
        return data
    
    def extract_all_data(self, article: str) -> Optional[Dict[str, Tuple[str, str]]]:
        """
        Полный процесс: найти паспорт и извлечь данные
        """
        passport_path = self.find_passport(article)
        
        if not passport_path:
            return None
        
        try:
            return self.extract_technical_data(passport_path)
        except Exception as e:
            raise Exception(f"Ошибка извлечения данных из {passport_path}: {str(e)}")
