import os
import re
import glob
import pandas as pd
import numbers_parser

DATA_DIR = os.path.join(os.path.dirname(__file__), '..', '..')

SALES_FILES = [
    'sales results 01.09.2025-30.11.2025.numbers',
    'sales results 01.12.2025-28.02.2026.numbers',
    'sales results 01.03.2026-31.03.2026.numbers',
]

SKIP_ARTICLES = {'77SK2024T'}


def load_sales() -> pd.DataFrame:
    dfs = []
    for fname in SALES_FILES:
        path = os.path.join(DATA_DIR, fname)
        doc = numbers_parser.Document(path)
        rows = list(doc.sheets[0].tables[0].rows())
        headers = [str(c.value) for c in rows[0]]
        data = [[c.value for c in row] for row in rows[1:]]
        dfs.append(pd.DataFrame(data, columns=headers))

    df = pd.concat(dfs, ignore_index=True)
    df['date'] = pd.to_datetime(df['Принят в обработку']).dt.normalize()
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(1)
    df['Оплачено покупателем'] = pd.to_numeric(df['Оплачено покупателем'], errors='coerce').fillna(0)
    df = df[~df['Артикул'].isin(SKIP_ARTICLES)]
    return df


def load_analytics() -> pd.DataFrame:
    pattern = os.path.join(DATA_DIR, 'analytics_report_*.xlsx')
    files = [f for f in glob.glob(pattern) if not re.search(r'\d{4}-\d{2}-\d{2}', os.path.basename(f))]

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
    return df
