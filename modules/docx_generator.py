"""
Модуль для генерации Word документов расчётов на прочность - ФИНАЛЬНАЯ ВЕРСИЯ
"""
import os
import shutil
import re
from typing import Dict, Tuple, Optional, List
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import win32com.client
from PIL import Image


class DOCXGenerator:
    """Класс для генерации документов Word"""
    
    def __init__(self, template_path: str, output_dir: str, config_manager, logger):
        """
        Инициализация генератора документов
        """
        self.template_path = template_path
        self.output_dir = output_dir
        self.config = config_manager
        self.logger = logger
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
    
    def generate_document(self, product_data: Dict, technical_data: Dict[str, Tuple[str, str]]) -> str:
        """
        Генерация документа для изделия
        """
        # Формируем имя выходного файла
        article = product_data['article'].replace('/', '-').replace('\\', '-')
        name = product_data['name'][:50]
        name = self._sanitize_filename(name)
        
        output_filename = f"{article}_{name}_РП.docx"
        output_path = os.path.join(self.output_dir, output_filename)
        
        # Копируем шаблон
        shutil.copy(self.template_path, output_path)
        
        # Открываем документ
        doc = Document(output_path)
        
        # Обрабатываем документ
        self._process_document(doc, product_data, technical_data)
        
        # Сохраняем
        doc.save(output_path)
        
        # Обновляем поля
        try:
            self._update_fields_com(output_path)
        except Exception as e:
            self.logger.log_warning(f"Не удалось обновить поля через COM: {str(e)}")
        
        return output_path
    
    def _process_document(self, doc: Document, product_data: Dict, technical_data: Dict):
        """
        Полная обработка документа
        """
        # 1. Подготовка данных
        replacements = self._prepare_replacements(product_data, technical_data)
        
        # 2. Замена текста в параграфах
        for paragraph in doc.paragraphs:
            self._replace_in_paragraph(paragraph, replacements)
        
        # 3. Замена текста в таблицах (включая штамп)
        self._update_tables(doc, replacements, product_data)
        
        # 4. Удаление ненужных параграфов (подписи к рисункам)
        self._remove_figure_captions(doc)
        
        # 5. Удаление изображений кроме основного
        self._remove_extra_images(doc)
        
        # 6. Вставка изображения изделия
        if product_data.get('image_path') and os.path.exists(product_data['image_path']):
            self._insert_main_image(doc, product_data['image_path'])
    
    def _prepare_replacements(self, product_data: Dict, technical_data: Dict) -> Dict[str, str]:
        """
        Подготовка всех замен
        """
        article = product_data['article']
        name = product_data['name']
        category = product_data['category']
        children_count = product_data['children_count']
        
        # Параметры
        mass_child = self.config.get_mass_child()  # 53.8 кг
        total_mass = children_count * mass_child
        
        # Расчёт сил
        g = 10  # м/с²
        Fh = 0.2 * total_mass * g
        Fz = total_mass * g
        
        # Родительный падеж категории
        category_gen = self._get_category_genitive(category)
        
        # Базовые замены
        replacements = {
            # Титульный лист и штамп
            'Змейка без песочницы арт.810152': f'{name} арт.{article}',
            'Змейка без песочницы': name,
            'арт.810152': f'арт.{article}',
            '810152': article,
            
            # В тексте
            'GA8808': article,
            'артикул GA8808': f'артикул {article}',
            'песочницы, артикул GA8808': f'{category_gen}, артикул {article}',
            'песочницы': category_gen,
            
            # Нагрузки от детей
            '10 детей': f'{children_count} детей',
            '32.5 кг': f'{mass_child:.1f} кг',
            '32,5 кг': f'{mass_child:.1f} кг',
            
            # Силы
            'Fh = 646,8 Н': f'Fh = {Fh:.1f} Н',
            'Fh = 646.8 Н': f'Fh = {Fh:.1f} Н',
            'Fz = 6468 Н': f'Fz = {Fz:.0f} Н',
            'Fz = 6468.0 Н': f'Fz = {Fz:.0f} Н',
        }
        
        return replacements
    
    @staticmethod
    def _get_category_genitive(category: str) -> str:
        """Категория в родительном падеже"""
        genitive_map = {
            'Домики': 'игрового домика',
            'Игровые комплексы': 'игрового комплекса',
            'Игровые элементы': 'игрового элемента',
            'Мини-беседки': 'мини-беседки',
            'Беседки': 'беседки',
            'Песочницы': 'песочницы'
        }
        return genitive_map.get(category, 'конструкции')
    
    @staticmethod
    def _replace_in_paragraph(paragraph, replacements: Dict[str, str]):
        """Замена текста в параграфе"""
        full_text = paragraph.text
        
        modified = False
        for old_text, new_text in replacements.items():
            if old_text in full_text:
                full_text = full_text.replace(old_text, new_text)
                modified = True
        
        if modified:
            if paragraph.runs:
                for run in paragraph.runs[1:]:
                    run.text = ''
                paragraph.runs[0].text = full_text
            else:
                paragraph.text = full_text
    
    def _update_tables(self, doc: Document, replacements: Dict[str, str], product_data: Dict):
        """
        Обновление таблиц (включая штамп)
        """
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, replacements)
    
    def _remove_figure_captions(self, doc: Document):
        """
        Удаление подписей к рисункам
        """
        paragraphs_to_remove = []
        
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text.strip()
            
            # Удаляем параграфы с подписями к рисункам
            if (text.startswith('Рис.') or 
                text.startswith('На Рис.') or
                'приведен' in text.lower() and 'рис' in text.lower()):
                
                # Сохраняем только "Рис. 1. Общий вид конструкции" - для основного изображения
                if not ('Общий вид' in text and 'Рис. 1' in text):
                    paragraphs_to_remove.append(paragraph)
        
        # Удаляем параграфы
        for paragraph in paragraphs_to_remove:
            p_element = paragraph._element
            p_element.getparent().remove(p_element)
    
    def _remove_extra_images(self, doc: Document):
        """
        Удаление всех изображений кроме основного
        """
        for paragraph in doc.paragraphs:
            # Оставляем изображение только если рядом есть "Рис. 1" или "Общий вид"
            keep_image = False
            if 'Рис. 1' in paragraph.text or 'Общий вид' in paragraph.text:
                keep_image = True
            
            # Удаляем изображения
            if not keep_image:
                for run in paragraph.runs:
                    drawings = run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
                    for drawing in drawings:
                        parent = drawing.getparent()
                        parent.remove(drawing)
    
    def _insert_main_image(self, doc: Document, image_path: str):
        """
        Вставка основного изображения
        """
        try:
            # Ищем параграф с "Рис. 1" или создаём после "ОБЩИЕ СВЕДЕНИЯ"
            target_paragraph = None
            
            for paragraph in doc.paragraphs:
                text = paragraph.text
                if 'Рис. 1' in text or ('Общий вид' in text and 'конструкци' in text):
                    target_paragraph = paragraph
                    break
            
            # Если не нашли, ищем после "ОБЩИЕ СВЕДЕНИЯ"
            if not target_paragraph:
                for i, paragraph in enumerate(doc.paragraphs):
                    if 'ОБЩИЕ СВЕДЕНИЯ' in paragraph.text:
                        if i + 3 < len(doc.paragraphs):
                            target_paragraph = doc.paragraphs[i + 3]
                        break
            
            if target_paragraph:
                # Открываем изображение
                img = Image.open(image_path)
                width, height = img.size
                
                # Масштабируем
                max_width = 6.0
                if width > height:
                    new_width = max_width
                    new_height = (height / width) * max_width
                else:
                    new_height = max_width
                    new_width = (width / height) * max_width
                
                # Удаляем существующие изображения в этом параграфе
                for run in target_paragraph.runs:
                    drawings = run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
                    for drawing in drawings:
                        parent = drawing.getparent()
                        parent.remove(drawing)
                
                # Очищаем текст если он короткий
                if len(target_paragraph.text.strip()) < 50:
                    target_paragraph.clear()
                
                # Добавляем изображение
                run = target_paragraph.add_run()
                run.add_picture(image_path, width=Inches(new_width))
                target_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
        except Exception as e:
            self.logger.log_warning(f"Ошибка вставки изображения: {str(e)}")
    
    @staticmethod
    def _update_fields_com(doc_path: str):
        """
        Обновление полей через COM
        """
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(doc_path))
            doc.Fields.Update()
            for toc in doc.TablesOfContents:
                toc.Update()
            doc.Save()
            doc.Close()
            word.Quit()
        except Exception as e:
            raise Exception(f"Ошибка обновления полей: {str(e)}")
    
    @staticmethod
    def _sanitize_filename(filename: str) -> str:
        """Очистка имени файла"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename.strip()
