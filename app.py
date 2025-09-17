from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from docx import Document
from datetime import datetime
import io
import os

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # В продакшене используйте настоящий секретный ключ

def replace_in_runs_preserve_formatting(paragraph, replacements):
    """
    Заменяет плейсхолдеры с максимальным сохранением форматирования
    """
    if not paragraph.runs:
        return
    
    # Собираем информацию о каждом символе и его форматировании
    char_data = []
    
    for run in paragraph.runs:
        for char in run.text:
            char_data.append({
                'char': char,
                'run': run,
                'font_name': run.font.name,
                'font_size': run.font.size,
                'bold': run.font.bold,
                'italic': run.font.italic,
                'underline': run.font.underline,
                'color': run.font.color.rgb if run.font.color.rgb else None
            })
    
    if not char_data:
        return
    
    # Получаем исходный текст
    original_text = "".join(cd['char'] for cd in char_data)
    
    # Выполняем замены
    new_text = original_text
    for placeholder, replacement in replacements.items():
        new_text = new_text.replace(placeholder, replacement)
    
    if new_text == original_text:
        return  # Нет изменений
    
    # Удаляем все существующие runs
    for run in paragraph.runs[:]:
        paragraph._element.remove(run._element)
    
    # Создаем новые runs с сохранением форматирования
    if new_text:
        # Находим первое место замены для определения базового форматирования
        base_formatting = None
        for placeholder in replacements.keys():
            if placeholder in original_text:
                pos = original_text.find(placeholder)
                if pos < len(char_data):
                    base_formatting = char_data[pos]
                    break
        
        if not base_formatting and char_data:
            base_formatting = char_data[0]
        
        # Создаем новый run с новым текстом
        new_run = paragraph.add_run(new_text)
        
        # Применяем форматирование
        if base_formatting:
            if base_formatting['font_name']:
                new_run.font.name = base_formatting['font_name']
            if base_formatting['font_size']:
                new_run.font.size = base_formatting['font_size']
            if base_formatting['bold'] is not None:
                new_run.font.bold = base_formatting['bold']
            if base_formatting['italic'] is not None:
                new_run.font.italic = base_formatting['italic']
            if base_formatting['underline'] is not None:
                new_run.font.underline = base_formatting['underline']
            if base_formatting['color']:
                new_run.font.color.rgb = base_formatting['color']


def process_document_in_memory(template_path, replacements):
    """
    Обрабатывает документ в памяти и возвращает байты обработанного документа
    """
    # Загружаем шаблон
    doc = Document(template_path)
    
    # Обрабатываем абзацы
    for p in doc.paragraphs:
        replace_in_runs_preserve_formatting(p, replacements)
    
    # Обрабатываем таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_runs_preserve_formatting(p, replacements)
    
    # Сохраняем документ в память
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    
    return doc_io


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Получаем данные из формы
        surname = request.form.get('surname', '').strip()
        name = request.form.get('name', '').strip()
        patronymic = request.form.get('patronymic', '').strip()
        inn = request.form.get('inn', '').strip()
        
        # Валидация
        if not all([surname, name, patronymic, inn]):
            flash('Все поля обязательны для заполнения', 'error')
            return redirect(url_for('index'))
        
        if len(inn) != 12 or not inn.isdigit():
            flash('ИНН должен содержать ровно 12 цифр', 'error')
            return redirect(url_for('index'))
        
        try:
            # Получаем текущую дату
            current_date = datetime.now()
            
            # Подготавливаем замены
            replacements = {
                "{Фамилия}": surname,
                "{Имя}": name,
                "{Отчество}": patronymic,
                "{ИНН}": inn,
                "{dd}.{mm}.{yyyy} г.": f"{current_date.day:02d}.{current_date.month:02d}.{current_date.year} г."
            }
            
            # Обрабатываем документ
            template_path = "42ea1332-1e5e-43db-90ac-9ec0b29f1bee.docx"
            if not os.path.exists(template_path):
                flash('Файл шаблона не найден', 'error')
                return redirect(url_for('index'))
            
            processed_doc = process_document_in_memory(template_path, replacements)
            
            # Формируем имя файла
            filename = f"bankruptcy_document_{surname}_{name}_{current_date.strftime('%Y%m%d')}.docx"
            
            # Отправляем файл для скачивания
            return send_file(
                processed_doc,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
        except Exception as e:
            flash(f'Ошибка при обработке документа: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080) 