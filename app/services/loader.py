import os
import re
import glob
import pandas as pd
import numbers_parser

DATA_DIR = os.path.join(os.path.dirname(__file__), '..', '..')

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
