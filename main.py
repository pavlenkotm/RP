"""
Основная программа для автоматической генерации расчётов на прочность
"""
import os
import sys
from typing import Dict

# Добавляем путь к модулям
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

from modules.config_manager import ConfigManager
from modules.logger import RPLogger
from modules.excel_reader import ExcelReader
from modules.pdf_parser import PDFParser
from modules.docx_generator import DOCXGenerator


class StrengthCalculationGenerator:
    """Главный класс программы генерации расчётов на прочность"""
    
    def __init__(self, config_path: str, texts_path: str):
        """
        Инициализация генератора
        
        Args:
            config_path: путь к config.json
            texts_path: путь к texts_by_category.json
        """
        # Загрузка конфигурации
        self.config = ConfigManager(config_path, texts_path)
        
        # Инициализация логгера
        log_path = self.config.get_path('log_file')
        self.logger = RPLogger(log_path)
        
        # Инициализация компонентов
        self.excel_reader = None
        self.pdf_parser = None
        self.docx_generator = None
    
    def initialize_components(self):
        """Инициализация всех компонентов программы"""
        try:
            # Excel Reader
            excel_path = self.config.get_path('excel')
            column_mapping = self.config.get_excel_columns()
            self.excel_reader = ExcelReader(excel_path, column_mapping)
            
            # PDF Parser
            passports_dir = self.config.get_path('passports')
            passport_pattern = self.config.get_passport_pattern()
            self.pdf_parser = PDFParser(passports_dir, passport_pattern)
            
            # DOCX Generator
            template_path = self.config.get_path('template_docx')
            output_dir = self.config.get_path('output_docs')
            self.docx_generator = DOCXGenerator(template_path, output_dir, self.config, self.logger)
            
            self.logger.log_info("Все компоненты инициализированы успешно")
            return True
            
        except Exception as e:
            self.logger.log_error('INIT', 'Система', 'ERR_INIT', str(e))
            return False
    
    def process_product(self, product: Dict) -> bool:
        """
        Обработка одного изделия
        
        Args:
            product: данные изделия из Excel
            
        Returns:
            True если успешно, False если ошибка
        """
        article = product['article']
        name = product['name']
        image_path = product['image_path']
        
        try:
            # 1. Проверка наличия картинки
            if not image_path or not os.path.exists(image_path):
                self.logger.log_error(
                    article, name, 'ERR_NO_IMAGE',
                    f"Файл изображения не найден: {image_path}",
                    image_path
                )
                return False
            
            # 2. Поиск паспорта
            passport_path = self.pdf_parser.find_passport(article)
            if not passport_path:
                self.logger.log_error(
                    article, name, 'ERR_NO_PASSPORT',
                    f"Паспорт не найден для артикула {article}"
                )
                return False
            
            # 3. Извлечение технических данных из паспорта
            try:
                technical_data = self.pdf_parser.extract_technical_data(passport_path)
                if not technical_data:
                    self.logger.log_warning(
                        f"Не удалось извлечь технические данные из паспорта {article}"
                    )
                    technical_data = {}
            except Exception as e:
                self.logger.log_error(
                    article, name, 'ERR_PDF_PARSE',
                    f"Ошибка парсинга PDF: {str(e)}"
                )
                return False
            
            # 4. Генерация документа
            try:
                output_path = self.docx_generator.generate_document(product, technical_data)
                self.logger.log_success(article, name, output_path)
                return True
            except Exception as e:
                self.logger.log_error(
                    article, name, 'ERR_TEMPLATE',
                    f"Ошибка генерации документа: {str(e)}"
                )
                return False
        
        except Exception as e:
            self.logger.log_error(
                article, name, 'ERR_UNKNOWN',
                f"Неизвестная ошибка: {str(e)}"
            )
            return False
    
    def run(self):
        """Запуск программы"""
        self.logger.log_start()
        
        # Инициализация компонентов
        if not self.initialize_components():
            self.logger.log_info("КРИТИЧЕСКАЯ ОШИБКА: Не удалось инициализировать компоненты")
            return
        
        # Загрузка данных из Excel
        try:
            self.logger.log_info(f"Загрузка данных из Excel: {self.config.get_path('excel')}")
            self.excel_reader.load_data()
            products = self.excel_reader.get_products()
            self.logger.log_info(f"Загружено изделий из Excel: {len(products)}")
        except Exception as e:
            self.logger.log_error('EXCEL', 'Система', 'ERR_EXCEL_READ', str(e))
            return
        
        # Фильтрация по категориям
        allowed_categories = self.config.get_categories()
        if allowed_categories:
            products = [p for p in products if p['category'] in allowed_categories]
            self.logger.log_info(f"После фильтрации по категориям осталось: {len(products)} изделий")
        
        # Обработка изделий
        self.logger.log_info("-" * 80)
        self.logger.log_info("НАЧАЛО ОБРАБОТКИ ИЗДЕЛИЙ")
        self.logger.log_info("-" * 80)
        
        for idx, product in enumerate(products, 1):
            self.logger.log_info(f"\n[{idx}/{len(products)}] Обработка: {product['article']} - {product['name']}")
            self.process_product(product)
        
        # Итоговая статистика
        self.logger.log_summary()


def main():
    """Главная функция программы"""
    # Определяем пути к конфигурации
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, 'config', 'config.json')
    texts_path = os.path.join(script_dir, 'config', 'texts_by_category.json')
    
    # Проверка наличия конфигурационных файлов
    if not os.path.exists(config_path):
        print(f"ОШИБКА: Файл конфигурации не найден: {config_path}")
        print("Создайте файл config.json в папке config/")
        return
    
    if not os.path.exists(texts_path):
        print(f"ОШИБКА: Файл текстов не найден: {texts_path}")
        print("Создайте файл texts_by_category.json в папке config/")
        return
    
    # Создание и запуск генератора
    try:
        generator = StrengthCalculationGenerator(config_path, texts_path)
        generator.run()
    except KeyboardInterrupt:
        print("\n\nПрограмма прервана пользователем")
    except Exception as e:
        print(f"\n\nКРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
