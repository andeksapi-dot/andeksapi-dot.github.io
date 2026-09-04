#!/usr/bin/env python3
"""
Парсер расписания БНТУ ФТУГ из Excel-файла.
Читает .xls, выдаёт JSON со структурой:
  groups -> days -> slots -> lessons
"""

import re
import json
import sys
from collections import OrderedDict

try:
    import xlrd
except ImportError:
    print("pip3 install xlrd")
    sys.exit(1)

DAYS_ORDER = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье']


def normalize_lesson_type(t):
    t = t.strip().lower().rstrip(')')
    if '//' in t:
        # комбинированный тип "лекционное//практическое занятие"
        return 'лекция/практика'
    if 'лекцион' in t:
        return 'лекция'
    if 'практическ' in t:
        return 'практика'
    if 'лаборатор' in t:
        return 'лабораторная'
    if 'оиз' in t:
        return 'ОИЗ'
    return None


def find_teachers(text):
    """Ищет преподавателей в тексте, возвращает список и остаток текста"""
    teachers = []
    rest = text
    
    # Паттерн: пр.Фамилия И.О., доц.Фамилия И.О., ст.пр.Фамилия И.О., проф.Фамилия И.О.
    # Также ст.пр.Фамилия (без инициалов)
    # Также "пр.Фамилия И.О, ст.пр.Фамилия И.О." (с запятой)
    pattern = r'(?:пр\.|доц\.|ст\.пр\.|проф\.)\s*[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?)?'
    
    matches = re.findall(pattern, rest)
    for m in matches:
        teachers.append(m.strip())
        rest = rest.replace(m, '', 1)
    
    # "вак." - вакансия
    if 'вак.' in rest:
        teachers.append('вакансия')
        rest = rest.replace('вак.', '')
    
    return teachers, rest


def find_auditorium(text):
    """Ищет аудиторию и корпус, возвращает (auditorium, building, rest)"""
    rest = text
    
    # Сначала чистим .0 от float xlrd (506.0 -> 506)
    rest = re.sub(r'(\d+)\.0(?:\s+к|\s|$)', r'\1 ', rest)
    
    # Паттерн 1: "а 522 к 8", "а 522 к.8" (с предлогом "а")
    m = re.search(r'(?:^|\s)а\s+(\d{2,4}[а-я]?(?:\s*[а-я])?)\s*\.?\s*к\s*\.?\s*(\d{1,2})(?!\d)', rest)
    if m:
        room = m.group(1).strip()
        building = m.group(2).strip()
        if len(room) < 7 and len(building) < 4:
            rest = rest.replace(m.group(0), ' ', 1)
            return room, f'к{building}', rest
    
    # Паттерн 2: "522 к 8", "522 к.8" (без предлога)
    m = re.search(r'(?<!\w)(\d{2,4}[а-я]?)\s*\.?\s*к\s*\.?\s*(\d{1,2})(?!\d)', rest)
    if m:
        room = m.group(1).strip()
        building = m.group(2).strip()
        if len(room) < 7 and len(building) < 4:
            rest = rest.replace(m.group(0), ' ', 1)
            return room, f'к{building}', rest
    
    # Паттерн 3: "а 508 А к 11" (с буквой корпуса после цифры)
    m = re.search(r'(?:^|\s)а\s+(\d{2,4})\s+([а-яА-Я])\s*к\s*\.?\s*(\d{1,2})(?!\d)', rest)
    if m:
        room = m.group(1).strip() + ' ' + m.group(2).strip()
        building = m.group(3).strip()
        rest = rest.replace(m.group(0), ' ', 1)
        return room, f'к{building}', rest
    
    # Паттерн 4: "205 б к 9" (цифры буква к цифры, без предлога)
    m = re.search(r'(?<!\w)(\d{2,4})\s+([а-я])\s*к\s*\.?\s*(\d{1,2})(?!\d)', rest)
    if m:
        room = m.group(1).strip() + ' ' + m.group(2).strip()
        building = m.group(3).strip()
        rest = rest.replace(m.group(0), ' ', 1)
        return room, f'к{building}', rest
    
    return '', '', rest


def find_week(text):
    """Ищет номер недели, возвращает (week, rest)"""
    rest = text
    m = re.search(r'(\d)\s*нед[\.\s]', rest)
    if m:
        week = int(m.group(1))
        rest = rest.replace(m.group(0), ' ', 1)
        return week, rest
    return 0, rest


def find_subgroup(text):
    """Ищет подгруппу в тексте ('1/2 группы'), возвращает (int, rest)"""
    rest = text
    m = re.search(r'(\d)\s*/\s*2\s*(?:группы|гр\.?|подгруппы|подгр\.?)', rest)
    if m:
        sg = int(m.group(1))
        rest = rest.replace(m.group(0), ' ', 1)
        return sg, rest
    m = re.search(r'(\d)\s*(?:подгруппа|подгр\.)', rest)
    if m:
        sg = int(m.group(1))
        rest = rest.replace(m.group(0), ' ', 1)
        return sg, rest
    return '', rest


def parse_cell_text(text):
    """Парсит текст ячейки на список занятий"""
    if not text or not text.strip():
        return []
    
    text = text.strip()
    
    # Предварительно извлекаем тип занятия из всего текста, ДО разбивки на недели —
    # иначе "(лекционное занятие)   1 нед.  ..." порежется на пустой блок с типом, который выбросится.
    preamble_type = None
    preamble_rest = text
    while True:
        m = re.match(r'^\s*\(([^)]+)\)\s*', preamble_rest)
        if not m: break
        t = normalize_lesson_type(m.group(1))
        if t:
            preamble_type = preamble_type or t
            preamble_rest = preamble_rest[m.end():]
        elif 'ОИЗ' in m.group(1):
            preamble_type = preamble_type or 'ОИЗ'
            preamble_rest = preamble_rest[m.end():]
        else:
            break
    
    # Физкультура
    if text.lower().startswith('физическая культура'):
        teachers, rest = find_teachers(text)
        return [{
            'type': 'физкультура',
            'week': 0,
            'subject': 'Физическая культура',
            'auditorium': '',
            'building': '',
            'teachers': teachers,
            'subgroup': '',
            'raw': text
        }]
    
    # Производственное обучение
    if 'производственное' in text.lower() and 'обучение' in text.lower():
        return [{
            'type': 'практика',
            'week': 0,
            'subject': 'Производственное обучение',
            'auditorium': '',
            'building': '',
            'teachers': [],
            'subgroup': '',
            'raw': text
        }]
    
    # Разделяем на блоки по неделям
    # Ищем "1 нед." и "2 нед." как разделители
    blocks = re.split(r'(?=(?:\([^)]*\))?\s*\d\s*нед[\.\s])', text)
    blocks = [b.strip() for b in blocks if b.strip()]
    
    if len(blocks) <= 1:
        # Пробуем разделить по "1 неделя", "2 неделя"
        blocks = re.split(r'(?=(?:\([^)]*\))?\s*\d\s*неделя)', text)
        blocks = [b.strip() for b in blocks if b.strip()]
    
    if not blocks:
        blocks = [text]
    
    results = []
    any_parsed = False
    for block in blocks:
        if not block.strip():
            continue
        # Если блок пустой/состоит только из пробелов (типа от " (лекционное занятие)  1 нед.")
        # и есть preamble_type — это маркер типа для следующего блока
        if not block.strip() and preamble_type:
            continue
        lesson = parse_single_block(block.strip())
        if not lesson:
            continue
        # Применяем тип из преамбулы, если в самом блоке не нашлось
        if preamble_type and not lesson.get('type'):
            lesson['type'] = preamble_type
        any_parsed = True
        # Мусор: нет названия (в т.ч. одинокие буквы/обрывки типа "Ф", "с") — не занятие
        subj = lesson['subject'].strip()
        if len(subj) <= 1:
            continue
        results.append(lesson)
    
    if results or any_parsed:
        return results
    return [{'type': 'unknown', 'week': 0, 'subject': text, 'auditorium': '', 'building': '', 'teachers': [], 'subgroup': '', 'raw': text}]


def parse_single_block(block):
    """Парсит один блок текста"""
    if not block:
        return None
    
    result = {
        'type': '',
        'week': 0,
        'subject': '',
        'auditorium': '',
        'building': '',
        'teachers': [],
        'subgroup': '',
        'raw': block
    }
    
    rest = block
    
    # 1. Тип занятия в скобках (может повторяться — снимаем все ведущие маркеры типа)
    while True:
        type_match = re.search(r'\(([^)]+)\)', rest)
        if not type_match:
            break
        type_text = type_match.group(1)
        lesson_type = normalize_lesson_type(type_text)
        if lesson_type:
            if not result['type']:
                result['type'] = lesson_type
            rest = rest.replace(type_match.group(0), ' ', 1).strip()
        elif 'ОИЗ' in type_text:
            if not result['type']:
                result['type'] = 'ОИЗ'
            rest = rest.replace(type_match.group(0), ' ', 1).strip()
        else:
            break
    
    # ОИЗ без скобок
    if not result['type']:
        m = re.search(r'\bОИЗ\b', rest)
        if m:
            # Проверяем, что это не часть названия
            result['type'] = 'ОИЗ'
            rest = rest.replace(m.group(0), ' ', 1).strip()
    
    # 2. Неделя
    result['week'], rest = find_week(rest)
    
    # Если не нашли неделю, ищем "1 неделя", "2 неделя"
    if not result['week']:
        m = re.search(r'(\d)\s*неделя', rest)
        if m:
            result['week'] = int(m.group(1))
            rest = rest.replace(m.group(0), ' ', 1).strip()
    
    # 3. Подгруппа
    result['subgroup'], rest = find_subgroup(rest)
    
    # 4. Аудитория
    result['auditorium'], result['building'], rest = find_auditorium(rest)
    
    # 5. Преподаватели
    result['teachers'], rest = find_teachers(rest)
    
    # 6. Оставшийся текст - название предмета
    rest = rest.strip()
    # Убираем даты "14/9, 28/9"
    rest = re.sub(r'\d{1,2}/\d{1,2}(?:,\s*\d{1,2}/\d{1,2})*', '', rest)
    # Убираем время "15.40, 17.45" / "9.55-10.40" (аудиторные числа не трогаем)
    rest = re.sub(r'(?<![\d.])\d{1,2}\.\d{2}(?:\s*[,\-–]\s*\d{1,2}\.\d{2})*', '', rest)
    # Убираем "с 29/9"
    rest = re.sub(r'с\s+\d{1,2}/\d{1,2}', '', rest)
    # Убираем одинокие "а" (предлоги, остатки от "аудитория")
    rest = re.sub(r'\s+а\s+', ' ', rest)  # "а" между словами
    rest = re.sub(r'\s+а$', '', rest)      # "а" в конце
    rest = re.sub(r'^а\s+', '', rest)      # "а" в начале
    # Одинокое "с" в начале — остаток от "с 5/10" после зачистки дат
    rest = re.sub(r'^[Сс]\s+', '', rest)
    # Убираем лишние пробелы
    rest = re.sub(r'\s+', ' ', rest).strip()
    # Убираем знаки препинания в начале/конце
    rest = rest.strip(' ,;.-')
    # Второй проход: после пунктуации могли обнажиться "а"/"с" на концах
    rest = re.sub(r'\s[АаСс]$', '', rest)   # "язык а" / "язык с"
    rest = re.sub(r'^[АаСс]\s+', '', rest)  # "с Физика"
    rest = re.sub(r'\s+', ' ', rest).strip()
    
    if rest:
        # Капитализация первой буквы
        result['subject'] = rest[0].upper() + rest[1:] if len(rest) > 1 else rest
    
    return result


def cell_str(value):
    """str() числа без '.0' (xlrd хранит числа как float)"""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def make_grid(sh, group_row=11):
    """Сетка значений листа с учётом объединённых ячеек:
    значение merged-ячейки копируется во все её клетки (совмещёнки).
    group_row — строка с номерами групп (после неё начинается зона занятий).
    """
    grid = []
    for r in range(sh.nrows):
        row = []
        for c in range(sh.ncols):
            cell = sh.cell(r, c)
            v = cell_str(cell.value).strip() if cell.ctype != xlrd.XL_CELL_EMPTY else ''
            v = re.sub(r'\s*\n\s*', ' ', v)
            row.append(v)
        grid.append(row)
    zone_start = group_row + 1
    for (r1, r2, c1, c2) in getattr(sh, 'merged_cells', []):
        # Копируем только мержи в области занятий: строки после шапки групп,
        # колонки >=2 (не дни/часы — они merged вертикально/горизонтально служебно)
        if r1 < zone_start or c1 < 2:
            continue
        v = grid[r1][c1] if r1 < len(grid) and c1 < len(grid[r1]) else ''
        if not v:
            continue
        for r in range(r1, min(r2, sh.nrows)):
            for c in range(c1, min(c2, sh.ncols)):
                grid[r][c] = v
    return grid


def collect_window(grid, start_row, end_row, col_from, col_to):
    """Текст прямоугольника сетки (строки start_row..end_row, колонки col_from..col_to).
    Подряд идущие одинаковые значения схлопываются — это дубли merged-значений.
    """
    texts = []
    for r in range(start_row, end_row):
        row = grid[r]
        for c in range(col_from, min(col_to, len(row))):
            v = row[c]
            if v and (not texts or texts[-1] != v):
                texts.append(v)
    return ' '.join(texts)


def find_group_row(sh):
    """Ищет строку с номерами групп (8-значные числа, как 10802126).
    Возвращает номер строки или None."""
    for r in range(min(sh.nrows, 30)):
        cnt = 0
        for c in range(sh.ncols):
            v = sh.cell_value(r, c)
            if isinstance(v, float) and v == int(v) and 10000000 <= v < 100000000:
                cnt += 1
            elif isinstance(v, str) and re.fullmatch(r'\d{8}', v.strip()):
                cnt += 1
        if cnt >= 2:
            return r
    return None


def parse_ftug_sheet(wb, sheet_name):
    """Парсит лист расписания (любой факультет БНТУ)."""
    sh = wb.sheet_by_name(sheet_name)
    group_row = find_group_row(sh)
    if group_row is None:
        print(f"  ⚠ нет строки групп на листе {sheet_name!r}")
        return {}
    grid = make_grid(sh, group_row=group_row)
    
    # Группы (каждая группа занимает 4 колонки)
    groups = []
    for c in range(sh.ncols):
        if len(grid) > group_row:
            m = re.match(r'(\d{8})', grid[group_row][c])
            if m:
                groups.append((m.group(1), c))
    
    if not groups:
        return {}
    
    # Merged-регионы в зоне занятий: строка после шапки групп, колонка >=2
    zone_start = group_row + 1
    merged = [(r1, r2, c1, c2) for (r1, r2, c1, c2) in getattr(sh, 'merged_cells', [])
              if r1 >= zone_start and c1 >= 2]
    
    def shared_groups(col_start, col_end, slot_start, next_slot):
        """Группы, с которыми блок колонок col_start..col_end делит merged-ячейку в слоте.
        Совмещёнка = merged-регион пересекает блок группы И выходит за его пределы."""
        shared = set()
        for (r1, r2, c1, c2) in merged:
            if r2 <= slot_start or r1 >= next_slot:
                continue
            if c2 <= col_start or c1 >= col_end:
                continue
            # Регион строго внутри блока — это просто занятие на всю ширину своей группы
            if c1 >= col_start and c2 <= col_end:
                continue
            for gname2, gcol in groups:
                if gcol == col_start:
                    continue  # свой блок — не совмещёнка
                if gcol < c2 and gcol + 4 > c1:
                    shared.add(gname2)
        return sorted(shared)
    
    # Дни недели — строки, где в колонке 0 название дня
    day_rows = [r for r in range(sh.nrows)
                if grid[r] and grid[r][0].strip().lower() in DAYS_ORDER]
    
    result = {g: {} for g, _ in groups}
    time_re = re.compile(r'\d{1,2}\.\d{2}\s*[-–]\s*\d{1,2}\.\d{2}')
    
    for di, day_row in enumerate(day_rows):
        day_name = grid[day_row][0].strip().lower()
        next_day_row = day_rows[di + 1] if di + 1 < len(day_rows) else sh.nrows
        
        # Слоты — строки с временем в колонке 1
        slot_rows = [day_row] + [r for r in range(day_row + 1, next_day_row)
                                 if len(grid[r]) > 1 and time_re.match(grid[r][1])]
        
        for si, slot_start in enumerate(slot_rows):
            next_slot = slot_rows[si + 1] if si + 1 < len(slot_rows) else next_day_row
            time_str = grid[slot_start][1] if len(grid[slot_start]) > 1 else ''
            
            for group_name, col_start in groups:
                col_end = col_start + 4
                # Две половины блока группы: левые 2 колонки и правые 2
                ltext = collect_window(grid, slot_start, next_slot, col_start, col_start + 2)
                rtext = collect_window(grid, slot_start, next_slot, col_start + 2, col_end)
                
                if ltext and rtext and ltext != rtext:
                    # Параллельные занятия: левая половина — подгруппа 1, правая — 2
                    lessons = []
                    for sub, wtext in ((1, ltext), (2, rtext)):
                        for les in parse_cell_text(wtext):
                            les['subgroup'] = sub
                            lessons.append(les)
                elif ltext and not rtext:
                    # Только 1-я подгруппа (2-я свободна)
                    lessons = []
                    for les in parse_cell_text(ltext):
                        les['subgroup'] = 1
                        lessons.append(les)
                elif rtext and not ltext:
                    # Только 2-я подгруппа (1-я свободна)
                    lessons = []
                    for les in parse_cell_text(rtext):
                        les['subgroup'] = 2
                        lessons.append(les)
                else:
                    full = collect_window(grid, slot_start, next_slot, col_start, col_end)
                    lessons = parse_cell_text(full) if full else []
                
                if lessons:
                    sw = shared_groups(col_start, col_end, slot_start, next_slot)
                    if sw:
                        for les in lessons:
                            les['shared_with'] = sw
                    result[group_name].setdefault(day_name, {})
                    result[group_name][day_name].setdefault(time_str, []).extend(lessons)
    
    return result


def main():
    import sys
    if len(sys.argv) < 3:
        print("Использование: parse_bntu.py <out.json> <xls1> [xls2 ...]")
        print("  Парсит один или несколько .xls (по файлу на курс) и сливает в общий JSON.")
        sys.exit(1)
    json_path = sys.argv[1]
    xls_paths = sys.argv[2:]
    course_labels = ['1 КУРС', '2 КУРС', '3 КУРС', '4 КУРС', '5 КУРС']
    
    all_data = {}
    for idx, xls_path in enumerate(xls_paths):
        course_label = course_labels[idx] if idx < len(course_labels) else f'{idx+1} КУРС'
        print(f"\n### Файл курса {idx+1}: {xls_path} (метка «{course_label}») ###")
        wb = xlrd.open_workbook(xls_path, formatting_info=True)
        for sheet_name in wb.sheet_names():
            data = parse_ftug_sheet(wb, sheet_name)
            if not data:
                continue
            # Листы типа «ФТУГ 1 КУРС» / «ФТУГ» (старшие курсы 1 файла) — оставляем как есть.
            # Листы без явной метки курса (типа «АТФ 1 КУРС» во 2-курсовом файле — а там реально 2 КУРС)
            # подменяем метку курса на ту, что соответствует файлу.
            new_key = sheet_name
            if re.search(r'\d\s*КУРС|\d\s*курс', sheet_name):
                # заменяем число в названии на course_label
                new_key = re.sub(r'\d+\s*КУРС', course_label, sheet_name, flags=re.IGNORECASE)
            else:
                # «ФТУГ» без курса — оставляем, но добавим пометку файла
                new_key = f"{sheet_name.strip()} ({course_label})"
            all_data[new_key] = data
            total = sum(len(lessons) for g in data for d in data[g] for t in data[g][d] for lessons in [data[g][d][t]])
            print(f"  {sheet_name!r} → ключ {new_key!r}: групп={len(data)}, занятий={total}")
    
    # Сохраняем
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    
    total_groups = sum(len(g) for s in all_data.values() for g in [s])
    total_lessons = sum(len(l) for s in all_data.values() for g in s.values() for d in g.values() for t in d.values() for l in t)
    print(f"\nJSON: {json_path}")
    print(f"  листов: {len(all_data)}, групп: {total_groups}, занятий: {total_lessons}")


if __name__ == '__main__':
    main()