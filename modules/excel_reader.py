"""
Модуль для чтения данных из Excel файла с каталогом изделий
"""
import pandas as pd
import os
from typing import List, Dict, Optional


class ExcelReader:
    """Класс для чтения данных из Excel файла"""
    
    def __init__(self, excel_path: str, column_mapping: Dict[str, str]):
        """
        Инициализация ридера Excel
        
        Args:
            excel_path: путь к Excel файлу
            column_mapping: маппинг столбцов (article, name, image_path, children_count)
        """
        self.excel_path = excel_path
        self.column_mapping = column_mapping
        self.df = None
    
    def load_data(self) -> bool:
        """
        Загрузка данных из Excel
        
        Returns:
            True если успешно, False если ошибка
        """
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Excel файл не найден: {self.excel_path}")
        
        try:
            # Читаем Excel, начиная со второй строки (индекс 1), если есть заголовок
            self.df = pd.read_excel(self.excel_path, engine='openpyxl')
            return True
        except Exception as e:
            raise Exception(f"Ошибка чтения Excel файла: {str(e)}")
    
    def get_products(self) -> List[Dict[str, any]]:
        """
        Получить список изделий из Excel
        
        Returns:
            Список словарей с данными изделий
        """
        if self.df is None:
            raise Exception("Данные не загружены. Вызовите load_data() сначала.")
        
        products = []
        
        # Получаем индексы столбцов из букв (A=0, B=1, C=2, D=3, E=4)
        col_article = self._column_letter_to_index(self.column_mapping['article'])
        col_name = self._column_letter_to_index(self.column_mapping['name'])
        col_image = self._column_letter_to_index(self.column_mapping['image_path'])
        col_children = self._column_letter_to_index(self.column_mapping['children_count'])
        
        for idx, row in self.df.iterrows():
            try:
                # Извлекаем данные по индексам столбцов
                article = self._safe_get_value(row, col_article)
                name = self._safe_get_value(row, col_name)
                image_path = self._safe_get_value(row, col_image)
                children_count = self._safe_get_value(row, col_children)
                
                # Пропускаем строки с пустым артикулом или наименованием
                if not article or not name:
                    continue
                
                # Определяем категорию (по умолчанию из наименования)
                category = self._determine_category(name)
                
                # Преобразуем количество детей в число
                try:
                    children_count = int(float(children_count)) if children_count else 1
                except (ValueError, TypeError):
                    children_count = 1
                
                product = {
                    'article': str(article).strip(),
                    'name': str(name).strip(),
                    'image_path': str(image_path).strip() if image_path else '',
                    'children_count': children_count,
                    'category': category,
                    'row_index': idx
                }
                
                products.append(product)
                
            except Exception as e:
                # Пропускаем строку с ошибкой
                continue
        
        return products
    
    @staticmethod
    def _column_letter_to_index(letter: str) -> int:
        """
        Преобразование буквы столбца в индекс
        
        Args:
            letter: буква столбца (A, B, C и т.д.)
            
        Returns:
            Индекс столбца (0-based)
        """
        letter = letter.upper().strip()
        index = 0
        for char in letter:
            index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1
    
    @staticmethod
    def _safe_get_value(row, col_index: int):
        """
        Безопасное получение значения из строки по индексу
        
        Args:
            row: строка DataFrame
            col_index: индекс столбца
            
        Returns:
            Значение ячейки или None
        """
        try:
            value = row.iloc[col_index]
            # Проверяем на NaN
            if pd.isna(value):
                return None
            return value
        except (IndexError, KeyError):
            return None
    
    @staticmethod
    def _determine_category(name: str) -> str:
        """
        Определение категории изделия по наименованию
        
        Args:
            name: наименование изделия
            
        Returns:
            Категория изделия
        """
        name_lower = name.lower()
        
        if 'домик' in name_lower:
            return 'Домики'
        elif 'комплекс' in name_lower:
            return 'Игровые комплексы'
        elif 'песочниц' in name_lower:
            return 'Песочницы'
        elif 'мини-беседк' in name_lower or 'миниbeседк' in name_lower:
            return 'Мини-беседки'
        elif 'беседк' in name_lower:
            return 'Беседки'
        else:
            return 'Игровые элементы'
    
    def get_product_count(self) -> int:
        """Получить количество изделий в Excel"""
        if self.df is None:
            return 0
        return len(self.df)
