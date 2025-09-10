from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Имена файлов
input_file = "title_page.docx"
output_file = "Лабораторная_работа_1_отчет.docx"

# Пути к изображениям
scheme_image = "схема АМ передатчика.png"
result_image = "Результат программы.png"

# Текст отчета с метками
report_text = """
[SECTION_START Goal]
Изучить основы устройства АМ сигналов и построить амплитудный модулятор (АМ-передатчик) в среде разработки GNU Radio Companion (GRC).
[SECTION_END Goal]

[SECTION_START Task]
1. Собрать схему АМ передатчика в среде разработки GRC согласно приведенному описанию.
2. Проанализировать изменение сигнала при увеличении параметра усилителя сигнала.
3. Ответить на контрольные вопросы.
[SECTION_END Task]

[SECTION_START Theory]
Амплитудная модуляция (АМ) – процесс изменения амплитуды высокочастотного несущего сигнала в соответствии с мгновенными значениями низкочастотного информационного (модулирующего) сигнала.

В России для радиовещания с АМ-модуляцией используется диапазон несущих частот 526,5–1606,5 кГц.

Структурная схема высокоуровневого АМ-передатчика включает в себя генератор несущей, буферный усилитель, предварительный усилитель звукового сигнала, голосовой процессор, усилитель модулятора и усилитель мощности, на котором и происходит модуляция.

Лабораторная установка: среда разработки GNU Radio Companion (GRC), использующая визуальное программирование для моделирования систем связи.
[SECTION_END Theory]

[SECTION_START Calculations]
4.1. Структурная схема исследуемого передатчика и описание блоков

Структурная схема собранного АМ-передатчика представлена на Рисунке 1.

[IMAGE_SCHEME]

Описание функций блоков:

• Options: Задает общие параметры потокового графа (название, автор, частота дискретизации).
• Variable (samp_rate): Определяет основную частоту дискретизации системы (768 кГц).
• Wav File Source: Источник сигнала. Считывает аудиоданные из WAV-файла.
• Repeat: Повторяет отсчеты аудиосигнала для согласования его низкой частоты дискретизации с высокой частотой дискретизации системы (768 кГц). Необходим для корректной работы последующих блоков.
• QT GUI Range: Графический регулятор для изменения коэффициента усиления в реальном времени.
• Multiply Const: Умножает входной сигнал на постоянный коэффициент (заданный QT GUI Range). Данный блок выполняет роль усилителя звукового сигнала, определяя глубину модуляции.
• Add Const: Добавляет к сигналу постоянное значение (+1). Сдвигает знакопеременный аудиосигнал в положительную область, так как амплитуда несущей не может быть отрицательной.
• Signal Source: Генерирует гармонический сигнал (несущую) частотой 48 кГц.
• Multiply: Ключевой блок модулятора. Перемножает модулирующий звуковой сигнал (с выхода Add Const) и несущую (с выхода Signal Source). Результатом является амплитудно-модулированное колебание.
• QT GUI Time Sink: Осциллограф. Визуализирует сигнал во временной области, позволяя наблюдать форму АМ-сигнала.

4.2. Графики изменения колебаний сигнала и выводы

Результат работы схемы представлен на Рисунке 2. На осциллографе (QT GUI Time Sink) наблюдается классический амплитудно-модулированный сигнал. Высокочастотное заполнение – это несущая частота 48 кГц. Огибающая (форма, описывающая пики колебаний) повторяет форму подаваемого аудиосигнала.

[IMAGE_RESULT]

Анализ изменения сигнала при увеличении параметра усилителя (Multiply Const):

1. Нулевое усиление (Constant = 0): Наблюдается чистая, немодулированная несущая с постоянной амплитудой. Вывод: модуляция не происходит.
2. Малое усиление (Constant = 0.3): Появляется слабо выраженная огибающая. Вывод: глубина модуляции мала.
3. Оптимальное усиление (Constant = 0.8 - 1.0): Четко видна огибающая, форма которой хорошо соответствует исходному аудиосигналу. Вывод: достигнута эффективная и качественная модуляция.
4. Слишком большое усиление (Constant > 1.2): Огибающая становится резкой, появляются искажения (овермодуляция). Вывод: происходит перемодуляция, приводящая к сильным искажениям.
[SECTION_END Calculations]

[SECTION_START Conclusions]
В ходе работы была успешно собрана и исследована модель АМ-передатчика. Экспериментально подтверждено, что коэффициент усиления в звуковом тракте напрямую определяет глубину модуляции. Найден диапазон значений для качественной модуляции и продемонстрированы негативные эффекты перемодуляции. Цель работы достигнута.
[SECTION_END Conclusions]

[SECTION_START Test]
1. В каком диапазоне несущих частот происходит радиовещание с АМ-модуляцией России?
Ответ: б) 526.5–1606.5 кГц

2. Частота дискретизации 768 кГц для обеспечения несущей частоты:
Ответ: а) 48 кГц 16 выборок за цикл (768 кГц / 48 кГц = 16)

3. Какой блок дает нам визуальное представление сигнала?
Ответ: а) QT GUI Time Sink

4. Зачем мы использовали блок Repeat?
Ответ: а) Чтобы повысить частоту дискретизации аудиовхода до определенной частоты
[SECTION_END Test]

[SECTION_START Bibliography]
1. Калач Г.П. Средства связи в системах управления автономными роботами: методические указания / Калач Г. П. – Москва: МИРЭА – Российский технологический университет, 2022. – 58 с.

2. OST 7.32-2017 Отчет о научно-исследовательской работе. Структура и правила оформления.
[SECTION_END Bibliography]
"""

def parse_sections(text):
    """Парсит текст и возвращает словарь с разделами"""
    sections = {}
    lines = text.strip().split('\n')
    current_section = None
    content = []
    
    for line in lines:
        if line.startswith('[SECTION_START'):
            current_section = line.split(' ')[1].replace(']', '')
            content = []
        elif line.startswith('[SECTION_END'):
            sections[current_section] = '\n'.join(content).strip()
            current_section = None
            content = []
        elif current_section is not None:
            content.append(line)
    
    return sections

def set_page_settings(doc):
    """Устанавливает настройки страницы согласно требованиям"""
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)

def add_heading(doc, text, level=1):
    """Добавляет заголовок с правильным стилем"""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.bold = True

def add_paragraph(doc, text, bold=False, italic=False):
    """Добавляет абзац с правильным форматированием"""
    if not text.strip():
        return
        
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.bold = bold
    run.italic = italic
    
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Cm(1.0)

def add_image_with_caption(doc, image_path, caption_text, width=Cm(12)):
    """Добавляет изображение с подписью по центру"""
    try:
        # Добавляем изображение
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        run.add_picture(image_path, width=width)
        
        # Добавляем подпись к изображению
        caption_paragraph = doc.add_paragraph()
        caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption_run = caption_paragraph.add_run(caption_text)
        caption_run.font.name = 'Times New Roman'
        caption_run.font.size = Pt(12)
        caption_run.italic = True
        
        # Добавляем пустую строку после изображения
        doc.add_paragraph()
        
    except FileNotFoundError:
        print(f"Файл изображения не найден: {image_path}")
        # Добавляем заглушку, если изображение не найдено
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(f"[Изображение: {image_path} не найдено]")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.italic = True

def process_text_with_images(doc, text, section_name):
    """Обрабатывает текст, заменяя метки изображений на реальные картинки"""
    lines = text.split('\n')
    
    for line in lines:
        if not line.strip():
            continue
            
        if line.strip() == '[IMAGE_SCHEME]':
            add_image_with_caption(doc, scheme_image, "Рисунок 1 – Структурная схема АМ-передатчика")
        elif line.strip() == '[IMAGE_RESULT]':
            add_image_with_caption(doc, result_image, "Рисунок 2 – Результат работы программы (АМ-сигнал)")
        else:
            add_paragraph(doc, line)

def setup_page_numbers(doc):
    """Настраивает нумерацию страниц"""
    section = doc.sections[0]
    footer = section.footer
    footer_paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_paragraph.add_run()
    
    fld_char = OxmlElement('w:fldChar')
    fld_char.set(qn('w:fldCharType'), 'begin')
    
    instr_text = OxmlElement('w:instrText')
    instr_text.set(qn('xml:space'), 'preserve')
    instr_text.text = "PAGE"
    
    fld_char2 = OxmlElement('w:fldChar')
    fld_char2.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fld_char)
    run._r.append(instr_text)
    run._r.append(fld_char2)

def main():
    # Парсим текст отчета на разделы
    sections = parse_sections(report_text)
    
    print("Найденные секции:", list(sections.keys()))
    
    # Открываем документ с титульным листом
    try:
        doc = Document(input_file)
    except FileNotFoundError:
        print(f"Файл {input_file} не найден. Создаем новый документ.")
        doc = Document()
    
    # Устанавливаем настройки страницы
    set_page_settings(doc)
    
    # Настраиваем нумерацию страниц
    setup_page_numbers(doc)
    
    # Добавляем разрыв страницы после титульного листа
    doc.add_page_break()
    
    # Добавляем разделы отчета
    if 'Goal' in sections:
        add_heading(doc, "1. Цель работы", level=1)
        process_text_with_images(doc, sections['Goal'], 'Goal')
    
    if 'Task' in sections:
        add_heading(doc, "2. Задача работы", level=1)
        process_text_with_images(doc, sections['Task'], 'Task')
    
    if 'Theory' in sections:
        add_heading(doc, "3. Теоретические сведения", level=1)
        process_text_with_images(doc, sections['Theory'], 'Theory')
    
    if 'Calculations' in sections:
        add_heading(doc, "4. Расчетно-графическая часть", level=1)
        process_text_with_images(doc, sections['Calculations'], 'Calculations')
    
    if 'Conclusions' in sections:
        add_heading(doc, "5. Выводы по работе", level=1)
        process_text_with_images(doc, sections['Conclusions'], 'Conclusions')
    
    if 'Test' in sections:
        add_heading(doc, "6. Ответы на контрольный тест", level=1)
        process_text_with_images(doc, sections['Test'], 'Test')
    
    if 'Bibliography' in sections:
        add_heading(doc, "7. Список используемой литературы", level=1)
        process_text_with_images(doc, sections['Bibliography'], 'Bibliography')
    
    # Сохраняем результат
    doc.save(output_file)
    print(f"Отчет успешно сохранен в файл: {output_file}")
    print(f"Использованные изображения: {scheme_image}, {result_image}")

if __name__ == "__main__":
    main()