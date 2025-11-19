"""
Модуль логирования для программы генерации расчётов на прочность
"""
import logging
import os
from datetime import datetime
from typing import Optional


class RPLogger:
    """Класс для логирования работы программы"""
    
    def __init__(self, log_file_path: str):
        """
        Инициализация логгера
        
        Args:
            log_file_path: путь к файлу лога
        """
        self.log_file_path = log_file_path
        self.stats = {
            'total': 0,
            'success': 0,
            'errors': {}
        }
        
        # Создаём директорию для лога, если её нет
        log_dir = os.path.dirname(log_file_path)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        
        # Настройка логгера
        self.logger = logging.getLogger('RPGenerator')
        self.logger.setLevel(logging.INFO)
        
        # Файловый обработчик
        file_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        
        # Консольный обработчик
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Формат логов
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Добавляем обработчики, если их ещё нет
        if not self.logger.handlers:
            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)
    
    def log_start(self):
        """Логирование начала работы программы"""
        self.logger.info("=" * 80)
        self.logger.info("НАЧАЛО РАБОТЫ ПРОГРАММЫ ГЕНЕРАЦИИ РАСЧЁТОВ НА ПРОЧНОСТЬ")
        self.logger.info("=" * 80)
    
    def log_success(self, article: str, name: str, output_path: str):
        """
        Логирование успешной обработки изделия
        
        Args:
            article: артикул изделия
            name: наименование изделия
            output_path: путь к созданному файлу
        """
        self.stats['total'] += 1
        self.stats['success'] += 1
        msg = f"OK | {article} | {name} | Файл сохранён: {output_path}"
        self.logger.info(msg)
    
    def log_error(self, article: str, name: str, error_code: str, 
                  error_msg: Optional[str] = None, image_path: Optional[str] = None):
        """
        Логирование ошибки при обработке изделия
        
        Args:
            article: артикул изделия
            name: наименование изделия
            error_code: код ошибки (ERR_NO_PASSPORT, ERR_NO_IMAGE и т.д.)
            error_msg: текстовое описание ошибки
            image_path: путь к картинке (если применимо)
        """
        self.stats['total'] += 1
        if error_code not in self.stats['errors']:
            self.stats['errors'][error_code] = 0
        self.stats['errors'][error_code] += 1
        
        msg = f"{error_code} | {article} | {name}"
        if image_path:
            msg += f" | Картинка: {image_path}"
        if error_msg:
            msg += f" | {error_msg}"
        
        self.logger.error(msg)
    
    def log_warning(self, message: str):
        """Логирование предупреждения"""
        self.logger.warning(message)
    
    def log_info(self, message: str):
        """Логирование информационного сообщения"""
        self.logger.info(message)
    
    def log_summary(self):
        """Вывод итоговой статистики"""
        self.logger.info("=" * 80)
        self.logger.info("ИТОГОВАЯ СТАТИСТИКА")
        self.logger.info("-" * 80)
        self.logger.info(f"Всего изделий обработано: {self.stats['total']}")
        self.logger.info(f"Успешно сформировано документов: {self.stats['success']}")
        
        error_count = self.stats['total'] - self.stats['success']
        self.logger.info(f"Изделий с ошибками: {error_count}")
        
        if self.stats['errors']:
            self.logger.info("-" * 80)
            self.logger.info("Распределение ошибок:")
            for error_code, count in self.stats['errors'].items():
                self.logger.info(f"  {error_code}: {count}")
        
        self.logger.info("-" * 80)
        self.logger.info(f"Лог-файл: {self.log_file_path}")
        self.logger.info("=" * 80)
        
        return self.stats
