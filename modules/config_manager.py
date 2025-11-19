"""
Модуль управления конфигурацией
"""
import json
import os
from typing import Dict, Any


class ConfigManager:
    """Класс для управления конфигурацией программы"""
    
    def __init__(self, config_path: str, texts_path: str):
        """
        Инициализация менеджера конфигурации
        
        Args:
            config_path: путь к файлу config.json
            texts_path: путь к файлу texts_by_category.json
        """
        self.config_path = config_path
        self.texts_path = texts_path
        self.config = self._load_json(config_path)
        self.texts = self._load_json(texts_path)
    
    @staticmethod
    def _load_json(file_path: str) -> Dict[str, Any]:
        """
        Загрузка JSON файла
        
        Args:
            file_path: путь к JSON файлу
            
        Returns:
            Словарь с данными из JSON
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Файл конфигурации не найден: {file_path}")
        
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def get_path(self, key: str) -> str:
        """Получить путь из конфигурации"""
        return self.config['paths'].get(key, '')
    
    def get_region(self) -> str:
        """Получить регион расчёта"""
        return self.config.get('region', 'Санкт-Петербург')
    
    def get_mass_child(self) -> float:
        """Получить массу одного ребёнка"""
        return self.config['loads'].get('mass_child', 53.8)
    
    def get_snow_load(self) -> Dict[str, Any]:
        """Получить параметры снеговой нагрузки"""
        return self.config['loads'].get('snow_load', {})
    
    def get_wind_load(self) -> Dict[str, Any]:
        """Получить параметры ветровой нагрузки"""
        return self.config['loads'].get('wind_load', {})
    
    def get_passport_pattern(self) -> str:
        """Получить паттерн для поиска паспорта"""
        return self.config.get('passport_pattern', '*{ART}*.pdf')
    
    def get_categories(self) -> list:
        """Получить список категорий изделий"""
        return self.config.get('categories', [])
    
    def get_excel_columns(self) -> Dict[str, str]:
        """Получить маппинг столбцов Excel"""
        return self.config.get('excel_columns', {})
    
    def get_category_texts(self, category: str) -> Dict[str, str]:
        """
        Получить тексты для категории изделия
        
        Args:
            category: категория изделия
            
        Returns:
            Словарь с текстами для категории
        """
        return self.texts.get(category, {
            'general_info': 'Объектом расчета является изделие',
            'construction_description': 'Конструкция представляет собой изделие',
            'conclusion': 'По результатам расчета установлено'
        })
    
    def is_debug_mode(self) -> bool:
        """Проверка режима отладки"""
        return self.config.get('debug_mode', False)
    
    def get_all_config(self) -> Dict[str, Any]:
        """Получить всю конфигурацию"""
        return self.config
