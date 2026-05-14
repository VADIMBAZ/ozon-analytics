import os
import re
import glob
import calendar
import pandas as pd
import numbers_parser

DATA_DIR = os.path.join(os.path.dirname(__file__), '..', '..')
UNIT_DIR = os.path.join(DATA_DIR, 'data', 'unit_economics')

# Файлы заказов подхватываются автоматически по маске. Канонический формат:
#   sales results <DD.MM.YYYY>-<DD.MM.YYYY>.<numbers|csv|xlsx>
# Имена с подчёркиванием/суффиксом во внутренней части = бэкап (см. фильтр ниже).
SALES_GLOB_EXTS = ('numbers', 'csv', 'xlsx')

SKIP_ARTICLES = {'77SK2024T'}


def _read_sales_file(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        return pd.read_csv(path, sep=';')
    if ext == '.xlsx':
        return pd.read_excel(path)
    doc = numbers_parser.Document(path)
    rows = list(doc.sheets[0].tables[0].rows())
    headers = [str(c.value) for c in rows[0]]
    data = [[c.value for c in row] for row in rows[1:]]
    return pd.DataFrame(data, columns=headers)


def _discover_sales_files() -> list[str]:
    """Находит все файлы заказов по маске, фильтруя бэкапы."""
    files = []
    for ext in SALES_GLOB_EXTS:
        files.extend(glob.glob(os.path.join(DATA_DIR, f'sales results *.{ext}')))
    # Фильтр бэкапов: .bak уже отсеивается через glob (нет в SALES_GLOB_EXTS).
    # Дополнительно отсекаем имена с " (1)", " копия" и т.п. — типичные дубли.
    files = [f for f in files
             if not any(suf in os.path.basename(f) for suf in (' (1)', ' (2)', ' копия', '_copy'))]
    return sorted(files)


def load_sales() -> pd.DataFrame:
    dfs = []
    files = _discover_sales_files()
    if not files:
        raise FileNotFoundError(
            f'Не найдено ни одного файла заказов в {DATA_DIR}. '
            'Ожидается «sales results <даты>.<numbers|csv|xlsx>».'
        )
    for path in files:
        dfs.append(_read_sales_file(path))

    df = pd.concat(dfs, ignore_index=True)
    df['date'] = pd.to_datetime(df['Принят в обработку']).dt.normalize()
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(1)
    df['Оплачено покупателем'] = pd.to_numeric(df['Оплачено покупателем'], errors='coerce').fillna(0)
    df = df[~df['Артикул'].isin(SKIP_ARTICLES)]
    return df


def load_analytics() -> pd.DataFrame:
    pattern = os.path.join(DATA_DIR, 'analytics_report_*.xlsx')
    # Канонический формат: analytics_report_<Месяц> <Год>.xlsx
    # Любое подчёркивание во «внутренней» части имени = бэкап/служебный файл (например, '_полный', '2026-03-26_12_44').
    files = []
    for f in glob.glob(pattern):
        inner = os.path.basename(f)[len('analytics_report_'):-len('.xlsx')]
        if '_' in inner:
            continue
        files.append(f)

    all_data = []
    for fpath in files:
        df_raw = pd.read_excel(fpath, header=None)
        period_str = str(df_raw.iloc[0, 0])
        dates = re.findall(r'(\d{2}\.\d{2}\.\d{4})', period_str)
        if len(dates) < 2:
            continue
        period_start = pd.to_datetime(dates[0], format='%d.%m.%Y')
        period_end = pd.to_datetime(dates[1], format='%d.%m.%Y')

        for i in range(13, len(df_raw)):
            row = df_raw.iloc[i]
            if str(row[0]) == 'nan':
                continue
            if str(row[5]) == 'nan':   # skip duplicate listings without model
                continue
            art = str(row[8])
            if art == 'nan' or art in SKIP_ARTICLES:
                continue

            days_str = str(row[27])
            days_no_stock = int(days_str.split(' из ')[0]) if 'из' in days_str else 0

            raw_stock = str(row[28])
            try:
                stock_end = int(float(raw_stock)) if raw_stock not in ('nan', '–') else None
            except (ValueError, TypeError):
                stock_end = None

            all_data.append({
                'period_start': period_start,
                'period_end': period_end,
                'month_label': period_start.strftime('%b %Y'),
                'Артикул': art,
                'Заказано': pd.to_numeric(row[17], errors='coerce'),
                'Доставлено': pd.to_numeric(row[19], errors='coerce'),
                'Отменено': pd.to_numeric(row[23], errors='coerce'),
                'дней_без_остатка': days_no_stock,
                'остаток_конец': stock_end,
            })

    df = pd.DataFrame(all_data).sort_values('period_start').reset_index(drop=True)

    # Регрессионная защита: один (артикул, период) должен встречаться ровно один раз.
    # Иначе значит, что glob подхватил бэкап/дубликат (исторический баг с '_полный.xlsx').
    dup_mask = df.duplicated(subset=['Артикул', 'period_start'], keep=False)
    if dup_mask.any():
        dup_rows = df[dup_mask].sort_values(['period_start', 'Артикул'])
        raise ValueError(
            'Дубли в analytics: один (артикул, период) встречается несколько раз. '
            'Скорее всего, в каталоге лежит бэкап (например, analytics_report_*_полный.xlsx). '
            f'Проблемные строки:\n{dup_rows[["period_start","Артикул"]].to_string(index=False)}'
        )
    return df


# ---------- Юнит-экономика ----------
# Канонический формат: Юнит-экономика_<DD.MM.YYYY>-<DD.MM.YYYY>.xlsx
# Только календарные месяцы (1.MM.YYYY до последнего числа того же месяца).
#
# Заголовки колонок Озон периодически меняет (с декабря 2025 добавилось 4 новых
# поля «Дополнительная обработка ОВХ», «Звёздные товары», «Платный бренд»,
# «Отзывы»). Поэтому индексы колонок ищем по тексту заголовка в строке 4,
# а не хардкодим. Один логический столбец может приходить под разными именами —
# тогда передаём список альтернатив.
UNIT_HEADER_ALIASES: dict[str, list[str]] = {
    'Артикул': ['Артикул'],
    'Цена': ['Текущая цена', 'Цена'],
    'Заказано': ['Заказано товаров, шт', 'Заказано'],
    'Доставлено': ['Доставлено товаров, шт', 'Доставлено'],
    'Возвращено': ['Возвращено товаров, шт', 'Возвращено'],
    'Выручка': ['Выручка'],
    'Баллы за скидки': ['Баллы за скидки'],
    'Программы партнёров': ['Программы партнёров'],
    'Вознаграждение Ozon': ['Вознаграждение Ozon'],
    'Эквайринг': ['Эквайринг'],
    'Обработка отправления': ['Обработка отправления'],
    'Логистика': ['Логистика'],
    'Доставка до ПВЗ': ['Доставка до места выдачи', 'Доставка до ПВЗ'],
    'Стоимость размещения': ['Стоимость размещения'],
    'Обработка возврата': ['Обработка возврата'],
    'Обратная логистика': ['Обратная логистика'],
    'Утилизация': ['Утилизация'],
    'Ошибки продавца': ['Обработка ошибок продавца', 'Операционные ошибки'],
    'Оплата за клик': ['Оплата за клик'],
    'Оплата за заказ': ['Оплата за заказ'],
    'Доля от продаж': ['Доля от продаж'],
    'Прибыль за шт': ['Прибыль за шт'],
    'Прибыль за период': ['Прибыль за период'],
}


def _build_header_index(header_row) -> dict[str, int]:
    """Из R4 строит {логическое имя → 0-based индекс колонки}. KeyError на пропавшее поле."""
    raw = {str(v).strip().replace('\n', ' ') if v is not None else '': i
           for i, v in enumerate(header_row)}
    out = {}
    for logical, aliases in UNIT_HEADER_ALIASES.items():
        for alias in aliases:
            if alias in raw:
                out[logical] = raw[alias]
                break
    return out


def _parse_unit_filename(path: str):
    """Извлекает (period_start, period_end, is_monthly) из имени файла или (None, None, False)."""
    name = os.path.basename(path)
    m = re.search(r'(\d{2}\.\d{2}\.\d{4})-(\d{2}\.\d{2}\.\d{4})', name)
    if not m:
        return None, None, False
    start = pd.to_datetime(m.group(1), format='%d.%m.%Y')
    end = pd.to_datetime(m.group(2), format='%d.%m.%Y')
    last_day = calendar.monthrange(start.year, start.month)[1]
    is_monthly = (
        start.day == 1
        and end.day == last_day
        and start.year == end.year
        and start.month == end.month
    )
    return start, end, is_monthly


def load_unit_economics() -> tuple[pd.DataFrame, list[str]]:
    """Парсит data/unit_economics/Юнит-экономика_*.xlsx.

    Возвращает (df, warnings). Файлы с произвольным диапазоном (не календарный месяц)
    пропускаются — их имя попадает в warnings.
    """
    warnings: list[str] = []
    if not os.path.isdir(UNIT_DIR):
        return pd.DataFrame(), warnings

    files = sorted(glob.glob(os.path.join(UNIT_DIR, 'Юнит-экономика_*.xlsx')))
    rows = []
    for fpath in files:
        start, end, is_monthly = _parse_unit_filename(fpath)
        if start is None:
            warnings.append(f'{os.path.basename(fpath)}: не удалось извлечь даты из имени')
            continue
        if not is_monthly:
            warnings.append(
                f'{os.path.basename(fpath)}: не календарный месяц ({start:%d.%m}–{end:%d.%m}), пропущено'
            )
            continue

        df_raw = pd.read_excel(fpath, header=None, sheet_name=0)
        # R2 может содержать маркер «экстраполировано из …» — для синтезированных
        # файлов (например, сшитый ноябрь из частичных выгрузок). UI покажет сноску.
        r2 = str(df_raw.iloc[1, 0]) if pd.notna(df_raw.iloc[1, 0]) else ''
        is_synth = 'экстраполи' in r2.lower()

        # Шапка занимает 4 строки. Колонки находим по тексту заголовка (R4),
        # а не по индексу — Озон периодически добавляет новые поля.
        header_idx = _build_header_index(df_raw.iloc[3])
        if 'Артикул' not in header_idx or 'Прибыль за период' not in header_idx:
            warnings.append(
                f'{os.path.basename(fpath)}: не нашёл обязательных колонок '
                f'(Артикул / Прибыль за период) в R4, пропущено'
            )
            continue

        art_col = header_idx['Артикул']
        for i in range(4, len(df_raw)):
            row = df_raw.iloc[i]
            art = str(row.iloc[art_col]) if pd.notna(row.iloc[art_col]) else 'nan'
            if art == 'nan' or art in SKIP_ARTICLES:
                continue
            rec = {
                'period_start': start,
                'period_end': end,
                'month_label': start.strftime('%b %Y'),
                'Артикул': art,
                'is_synthesized': is_synth,
                'source_note': r2 if is_synth else '',
            }
            for logical, col_idx in header_idx.items():
                if logical == 'Артикул':
                    continue
                rec[logical] = pd.to_numeric(row.iloc[col_idx], errors='coerce')
            rows.append(rec)

    if not rows:
        return pd.DataFrame(), warnings

    df = pd.DataFrame(rows).sort_values(['period_start', 'Артикул']).reset_index(drop=True)

    # Регрессионная защита: один (артикул, период) — одна строка
    dup_mask = df.duplicated(subset=['Артикул', 'period_start'], keep=False)
    if dup_mask.any():
        dup_rows = df[dup_mask][['period_start', 'Артикул']]
        raise ValueError(
            'Дубли в unit_economics: один (артикул, период) встречается несколько раз. '
            f'Проверь содержимое {UNIT_DIR}.\n{dup_rows.to_string(index=False)}'
        )
    return df, warnings
