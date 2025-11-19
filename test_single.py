"""
Тестовый скрипт для проверки работы программы на одном изделии
"""
import os
import sys

# Добавляем путь к модулям
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

from modules.config_manager import ConfigManager
from modules.logger import RPLogger
from modules.excel_reader import ExcelReader
from modules.pdf_parser import PDFParser
from modules.docx_generator import DOCXGenerator


def test_single_product():
    """Тестирование на одном изделии"""
    
    # Пути к конфигурации
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, 'config', 'config.json')
    texts_path = os.path.join(script_dir, 'config', 'texts_by_category.json')
    
    print("=" * 80)
    print("ТЕСТОВЫЙ ЗАПУСК - ОБРАБОТКА ОДНОГО ИЗДЕЛИЯ")
    print("=" * 80)
    
    # Инициализация компонентов
    try:
        config = ConfigManager(config_path, texts_path)
        print("✓ Конфигурация загружена")
    except Exception as e:
        print(f"✗ Ошибка загрузки конфигурации: {e}")
        return
    
    try:
        log_path = config.get_path('log_file').replace('.txt', '_test.txt')
        logger = RPLogger(log_path)
        print(f"✓ Логгер инициализирован: {log_path}")
    except Exception as e:
        print(f"✗ Ошибка инициализации логгера: {e}")
        return
    
    # Excel Reader
    try:
        excel_path = config.get_path('excel')
        column_mapping = config.get_excel_columns()
        excel_reader = ExcelReader(excel_path, column_mapping)
        excel_reader.load_data()
        print(f"✓ Excel загружен: {excel_path}")
    except Exception as e:
        print(f"✗ Ошибка загрузки Excel: {e}")
        return
    
    # Получаем список изделий
    try:
        products = excel_reader.get_products()
        print(f"✓ Найдено изделий в Excel: {len(products)}")
        
        if not products:
            print("✗ Нет изделий для обработки")
            return
        
        # Берём первое изделие
        test_product = products[0]
        print("\n" + "-" * 80)
        print("ТЕСТОВОЕ ИЗДЕЛИЕ:")
        print(f"  Артикул: {test_product['article']}")
        print(f"  Наименование: {test_product['name']}")
        print(f"  Категория: {test_product['category']}")
        print(f"  Количество детей: {test_product['children_count']}")
        print(f"  Путь к изображению: {test_product['image_path']}")
        print("-" * 80)
        
    except Exception as e:
        print(f"✗ Ошибка чтения изделий: {e}")
        return
    
    # PDF Parser
    try:
        passports_dir = config.get_path('passports')
        passport_pattern = config.get_passport_pattern()
        pdf_parser = PDFParser(passports_dir, passport_pattern)
        print(f"\n✓ PDF парсер инициализирован")
    except Exception as e:
        print(f"✗ Ошибка инициализации PDF парсера: {e}")
        return
    
    # Поиск паспорта
    try:
        passport_path = pdf_parser.find_passport(test_product['article'])
        if passport_path:
            print(f"✓ Паспорт найден: {passport_path}")
        else:
            print(f"✗ Паспорт НЕ найден для артикула: {test_product['article']}")
            print(f"  Искали в: {passports_dir}")
            print(f"  По шаблону: {passport_pattern.replace('{ART}', test_product['article'])}")
            return
    except Exception as e:
        print(f"✗ Ошибка поиска паспорта: {e}")
        return
    
    # Извлечение данных из паспорта
    try:
        technical_data = pdf_parser.extract_technical_data(passport_path)
        print(f"✓ Извлечено технических параметров: {len(technical_data)}")
        if technical_data:
            print("\n  Технические параметры:")
            for param, (value, unit) in list(technical_data.items())[:5]:
                print(f"    - {param}: {value} {unit}")
            if len(technical_data) > 5:
                print(f"    ... и ещё {len(technical_data) - 5} параметров")
    except Exception as e:
        print(f"✗ Ошибка извлечения данных из PDF: {e}")
        technical_data = {}
    
    # Проверка изображения
    if os.path.exists(test_product['image_path']):
        print(f"\n✓ Изображение найдено: {test_product['image_path']}")
    else:
        print(f"\n✗ Изображение НЕ найдено: {test_product['image_path']}")
    
    # DOCX Generator
    try:
        template_path = config.get_path('template_docx')
        output_dir = config.get_path('output_docs')
        docx_generator = DOCXGenerator(template_path, output_dir, config, logger)
        print(f"\n✓ DOCX генератор инициализирован")
        print(f"  Шаблон: {template_path}")
        print(f"  Выходная папка: {output_dir}")
    except Exception as e:
        print(f"✗ Ошибка инициализации DOCX генератора: {e}")
        return
    
    # Генерация документа
    print("\n" + "=" * 80)
    print("ГЕНЕРАЦИЯ ДОКУМЕНТА...")
    print("=" * 80)
    
    try:
        output_path = docx_generator.generate_document(test_product, technical_data)
        print(f"\n✓✓✓ УСПЕХ! ✓✓✓")
        print(f"\nДокумент создан: {output_path}")
        print(f"\nОткройте файл в Word для проверки результата.")
    except Exception as e:
        print(f"\n✗✗✗ ОШИБКА! ✗✗✗")
        print(f"\nОшибка генерации документа: {e}")
        import traceback
        traceback.print_exc()
        return
    
    print("\n" + "=" * 80)
    print("ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
    print("=" * 80)


if __name__ == "__main__":
    try:
        test_single_product()
    except KeyboardInterrupt:
        print("\n\nТестирование прервано пользователем")
    except Exception as e:
        print(f"\n\nКРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        import traceback
        traceback.print_exc()
