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
        # Получаем основные данные из формы
        surname = request.form.get('surname', '').strip()
        name = request.form.get('name', '').strip()
        patronymic = request.form.get('patronymic', '').strip()
        passport_data = request.form.get('passport_data', '').strip()
        registered_address = request.form.get('registered_address', '').strip()
        inn = request.form.get('inn', '').strip()
        snils = request.form.get('snils', '').strip()
        debt_amount_digits = request.form.get('debt_amount_digits', '').strip()
        debt_amount_words = request.form.get('debt_amount_words', '').strip()
        
        # Валидация основных полей
        required_fields = [surname, name, patronymic, passport_data, registered_address, inn, snils, debt_amount_digits, debt_amount_words]
        if not all(required_fields):
            flash('Все основные поля обязательны для заполнения', 'error')
            return redirect(url_for('index'))
        
        if len(inn) != 12 or not inn.isdigit():
            flash('ИНН должен содержать ровно 12 цифр', 'error')
            return redirect(url_for('index'))
        
        # Получаем данные о кредиторах
        creditors = []
        creditor_num = 1
        while True:
            creditor_name = request.form.get(f'creditor_name_{creditor_num}', '').strip()
            creditor_address = request.form.get(f'creditor_address_{creditor_num}', '').strip()
            
            if not creditor_name and not creditor_address:
                break
            
            if not creditor_name or not creditor_address:
                flash(f'Заполните все поля для кредитора {creditor_num}', 'error')
                return redirect(url_for('index'))
            
            creditors.append({
                'name': creditor_name,
                'address': creditor_address
            })
            creditor_num += 1
        
        if not creditors:
            flash('Необходимо указать хотя бы одного кредитора', 'error')
            return redirect(url_for('index'))
        
        try:
            # Получаем текущую дату
            current_date = datetime.now()
            
            # Получаем первые буквы имени и отчества
            first_letter_name = name[0].upper() if name else ''
            first_letter_patronymic = patronymic[0].upper() if patronymic else ''
            
            # Получаем название месяца на русском языке
            months_ru = [
                '', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
                'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
            ]
            month_name = months_ru[current_date.month]
            
            # Подготавливаем основные замены
            replacements = {
                "{Фамилия}": surname,
                "{Имя}": name,
                "{Отчество}": patronymic,
                "{паспортные данные}": passport_data,
                "{Зарегистрирован по адресу}": registered_address,
                "{ИНН}": inn,
                "{СНИЛС}": snils,
                "{сумма долга цифрами}": debt_amount_digits,
                "{сумма долга буквами}": debt_amount_words,
                "{dd}.{mm}.{yyyy} г.": f"{current_date.day:02d}.{current_date.month:02d}.{current_date.year} г.",
                "{число месяца}": str(current_date.day),
                "{месяц}": month_name,
                "{год}": str(current_date.year),
                "{Первая буква имени}": first_letter_name,
                "{первая буква отчества}": first_letter_patronymic
            }
            
            # Добавляем замены для кредиторов
            # Для первого кредитора используем общие плейсхолдеры
            if creditors:
                replacements["{Наименование кредитора}"] = creditors[0]['name']
                replacements["{Почтовый индекс и адрес}"] = creditors[0]['address']
            
            # Добавляем замены для всех кредиторов с номерами
            for i, creditor in enumerate(creditors, 1):
                replacements[f"{{Кредитор {i}}}"] = f"Кредитор {i}"
                replacements[f"{{Кредитор n}}"] = f"Кредитор {i}" if i == 1 else replacements.get(f"{{Кредитор n}}", f"Кредитор {i}")
                if i > 1:
                    replacements[f"{{Кредитор n+{i-1}}}"] = f"Кредитор {i}"
            
            # Обрабатываем документ
            template_path = "zayav.docx"
            if not os.path.exists(template_path):
                flash('Файл шаблона не найден', 'error')
                return redirect(url_for('index'))
            
            processed_doc = process_document_in_memory(template_path, replacements)
            
            # Формируем имя файла
            filename = f"bankruptcy_application_{surname}_{name}_{current_date.strftime('%Y%m%d')}.docx"
            
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