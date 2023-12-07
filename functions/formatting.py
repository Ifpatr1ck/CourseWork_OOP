from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
def indent_img(paragraphs):
    for paragraph in paragraphs:
        comment = paragraph.add_comment(f'Проверьте отступы у абзацев "Рисунок № -"')
        comment.author = 'bot'
        break
def check_name_font(expected_value, paragraphs):
    # чекает шрифт

    # проходит по всем параграфам
    for paragraph in paragraphs:
        # проходит по тексту в параграфе
        for run in paragraph.runs:
            # убираем пустые абзацы
            if paragraph.text != "":
                # шрифт Times New Roman он принимает за None, т.к. это стиль абзаца (др. шрифты, такие как Calibri, он тоже принимает :( )
                if run.font.name is not None:
                    # на всякий случай сделал еще это
                    if run.font.name != expected_value:
                        # написание коммента
                        comment = paragraph.add_comment('Шрифт должен быть Times New Roman!')
                        comment.author = 'bot'
def check_size_font(expected_value, paragraphs):
    # проходит по всем параграфам
    for paragraph in paragraphs:
        # проходит по тексту в параграфе
        for run in paragraph.runs:
            # убираем пустые абзацы
            if paragraph.text != "":
                # размер 14пт он принимает за None, т.к. это стиль абзаца
                if run.font.size is not None:
                    # на всякий случай сделал еще это
                    if run.font.size != expected_value:
                        # написание коммента
                        comment = paragraph.add_comment('Размер шрифта должен быть 14пт!')
                        comment.author = 'bot'
def check_first_line_indent(expected_value_first_line, paragraphs):
    # проверка отсутпа первого абзаца (вот тут трудности)
    # проходит по всем параграфам
    for paragraph in paragraphs:
        # переменная форматирования для удобства
        formatting = paragraph.paragraph_format
        # переменная отступа первого абзаца
        indent = formatting.first_line_indent
        # убираем пустые абзацы
        if paragraph.text != "":
            # убираем отступы, которые не считываются или их нет
            if indent is not None:
                # если отступ не равен введённому значению и выравнивание не по центру
                if round(indent.cm,
                         2) != expected_value_first_line and paragraph.alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
                    # написание коммента
                    comment = paragraph.add_comment(
                        f'Отступ первой строки абзаца должен быть {expected_value_first_line}см!')
                    comment.author = 'bot'
                # если отступ не равен 0 и выравнивание по центру
                elif round(indent.cm, 2) != 0.0 and paragraph.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
                    # написание коммента
                    comment = paragraph.add_comment(f'Отступ строки должен быть 0см!')
                    comment.author = 'bot'
def check_line_spacing(expected_value_line_spacing, paragraphs):
    # межстрочный интервал, обычно он комментит по 3-6 строк на одну строку
    # проходит по всем параграфам
    for paragraph in paragraphs:
        # переменная форматирования для удобства
        formatting = paragraph.paragraph_format
        # переменная отступа межстрочный
        spacing = formatting.line_spacing
        # убираем пустые абзацы
        if paragraph.text != "":
            # убираем отступы, которые не считываются или их нет
            if spacing is not None:
                # если отступ не равен введенному значению
                if round(spacing, 1) != expected_value_line_spacing:
                    # написание коммента
                    comment = paragraph.add_comment(
                        f'Межстрочный интервал должен быть {expected_value_line_spacing}см!')
                    comment.author = 'bot'
def check_heading(expected_value_first_line, paragraphs):
    # проверк азаголовка (надо улучшить), потому что за заголов он принимает все, чо пишется в начале жирным
    # проходит по всем параграфам
    for paragraph in paragraphs:
        # проходит по тексту в параграфе
        for run in paragraph.runs:
            # если текст в строчном не равен тексту написанный в документе (ЗАГЛАВНЫЕ) и он жирный и он не выравнен по центру
            if run.text.lower() != run.text and run.bold and paragraph.alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
                # написание коммента
                comment = paragraph.add_comment(f'Выравнивание должно быть по центру!')
                comment.author = 'bot'
