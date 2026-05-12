import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from services.loader import load_sales, load_analytics

COST_PER_UNIT = 25  # себестоимость, ₽/шт

st.set_page_config(page_title='Ozon Аналитика', layout='wide', page_icon='📦')

st.markdown("""
<style>
[data-testid="stSidebar"] { min-width: 260px; max-width: 260px; }
[data-testid="stSidebar"] button {
    white-space: nowrap;
    font-size: 13px;
    padding: 2px 10px;
    border-radius: 20px;
}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:has(> [data-testid="stCheckbox"]) {
    margin-top: -10px;
    margin-bottom: -10px;
}
</style>
""", unsafe_allow_html=True)

# ── Данные ──────────────────────────────────────────────────────────────────

@st.cache_data
def get_data():
    return load_sales(), load_analytics()

sales, analytics = get_data()

# ── Сайдбар ─────────────────────────────────────────────────────────────────

def _img_to_base64(path):
    import base64
    with open(path, 'rb') as f:
        return f'data:image/png;base64,{base64.b64encode(f.read()).decode()}'

ARTICLE_IMAGES = {
    '01SK2024': 'https://ir.ozone.ru/s3/multimedia-1-o/wc200/7558717596.jpg',
    '02SK2024': 'https://ir.ozone.ru/s3/multimedia-1-h/wc200/7790242049.jpg',
    '03SK2024': 'https://ir.ozone.ru/s3/multimedia-1-9/wc200/7558720029.jpg',
    '04SK2024': 'https://ir.ozone.ru/s3/multimedia-1-9/wc200/7790244741.jpg',
    '05SK2024': 'https://ir.ozone.ru/s3/multimedia-1-i/wc200/7790269914.jpg',
    '07SK2024': 'https://ir.ozone.ru/s3/multimedia-1-t/wc200/7790268917.jpg',
    '08SK2024': 'https://ir.ozone.ru/s3/multimedia-1-l/wc200/7790226033.jpg',
    '09SK2024': 'https://ir.ozone.ru/s3/multimedia-1-d/wc200/7790279377.jpg',
    '10SK2024': 'https://ir.ozone.ru/s3/multimedia-1-f/wc200/7558724715.jpg',
    '11SK2024': 'https://ir.ozone.ru/s3/multimedia-1-e/wc200/7790227898.jpg',
    '06SK2024': _img_to_base64(os.path.join(os.path.dirname(__file__), 'images', '06SK2024.png')),
    '14SK2024': _img_to_base64(os.path.join(os.path.dirname(__file__), 'images', '14SK2024.png')),
}

_delivered = sales[sales['Статус'] == 'Доставлен']
_sales_rank = _delivered.groupby('Артикул')['Количество'].sum().sort_values(ascending=False)
all_articles = [a for a in _sales_rank.index if a in sales['Артикул'].dropna().unique()]

# Инициализация состояния
for art in all_articles:
    if f'art_{art}' not in st.session_state:
        st.session_state[f'art_{art}'] = True

def reset_all():
    for art in all_articles:
        st.session_state[f'art_{art}'] = False

def select_all():
    for art in all_articles:
        st.session_state[f'art_{art}'] = True

btn_col1, btn_col2 = st.sidebar.columns(2)
btn_col1.button('✅ Все', on_click=select_all, use_container_width=True)
btn_col2.button('✖ Сброс', on_click=reset_all, use_container_width=True)

selected = []
for art in all_articles:
    img_url = ARTICLE_IMAGES.get(art)
    col_img, col_check = st.sidebar.columns([1, 2])
    if img_url:
        col_img.image(img_url, width=40)
    else:
        col_img.markdown(
            '<div style="width:40px;height:40px;background:#e0e0e0;border-radius:6px;'
            'display:flex;align-items:center;justify-content:center;'
            'font-size:10px;color:#999">фото</div>',
            unsafe_allow_html=True,
        )
    checked = col_check.checkbox(art, key=f'art_{art}')
    if checked:
        selected.append(art)

count = len(selected)
total = len(all_articles)
st.sidebar.caption(f'Выбрано: {count} из {total}')
st.sidebar.divider()
granularity = st.sidebar.radio('Гранулярность', ['Неделя', 'Месяц'], horizontal=True)

if not selected:
    st.warning('Выберите хотя бы один артикул.')
    st.stop()

st.markdown('<h1 style="text-align:center">📦 Спутник Ключи — Аналитика Ozon</h1>', unsafe_allow_html=True)

# ── KPI ─────────────────────────────────────────────────────────────────────

delivered_all = sales[sales['Статус'] == 'Доставлен']
cancelled_all = sales[sales['Статус'] == 'Отменён']

delivered_sel = delivered_all[delivered_all['Артикул'].isin(selected)]
cancelled_sel = cancelled_all[cancelled_all['Артикул'].isin(selected)]
total_orders = sales[sales['Артикул'].isin(selected)]['Количество'].sum()
revenue = delivered_sel['Оплачено покупателем'].sum()

# Дельты к прошлому месяцу
_sel_sales = sales[sales['Артикул'].isin(selected)].copy()
_sel_sales['_month'] = _sel_sales['date'].dt.to_period('M')
_months_sorted = sorted(_sel_sales['_month'].dropna().unique())
if len(_months_sorted) >= 2:
    _last_m, _prev_m = _months_sorted[-1], _months_sorted[-2]
    _cur = _sel_sales[_sel_sales['_month'] == _last_m]
    _prv = _sel_sales[_sel_sales['_month'] == _prev_m]
    _cur_dlv = _cur[_cur['Статус'] == 'Доставлен']
    _prv_dlv = _prv[_prv['Статус'] == 'Доставлен']
    _cur_canc = _cur[_cur['Статус'] == 'Отменён']
    _prv_canc = _prv[_prv['Статус'] == 'Отменён']
    _d_orders = int(_cur['Количество'].sum() - _prv['Количество'].sum())
    _d_dlv = int(_cur_dlv['Количество'].sum() - _prv_dlv['Количество'].sum())
    _d_canc = int(_cur_canc['Количество'].sum() - _prv_canc['Количество'].sum())
    _d_rev = _cur_dlv['Оплачено покупателем'].sum() - _prv_dlv['Оплачено покупателем'].sum()
    _d_orders_s = f"{_d_orders:+,}".replace(',', ' ')
    _d_dlv_s = f"{_d_dlv:+,}".replace(',', ' ')
    _d_canc_s = f"{_d_canc:+,}".replace(',', ' ')
    _d_rev_s = f"{_d_rev:+,.0f} ₽".replace(',', ' ')
else:
    _d_orders_s = _d_dlv_s = _d_canc_s = _d_rev_s = None

col1, col2, col3, col4 = st.columns(4)
col1.metric('Всего заказано', f"{int(total_orders):,}".replace(',', ' '), delta=_d_orders_s)
col2.metric('Доставлено', f"{int(delivered_sel['Количество'].sum()):,}".replace(',', ' '), delta=_d_dlv_s)
col3.metric('Отменено', f"{int(cancelled_sel['Количество'].sum()):,}".replace(',', ' '), delta=_d_canc_s, delta_color='inverse')
col4.metric('Выручка (доставл.)', f"{revenue:,.0f} ₽".replace(',', ' '), delta=_d_rev_s)

st.divider()

# ── Аналитические комментарии ────────────────────────────────────────────────

def _fmt(n):
    return f"{int(n):,}".replace(',', '\u202f')

def build_insights(sales_df, analytics_df, selected_arts):
    insights = []
    an = analytics_df[analytics_df['Артикул'].isin(selected_arts)].copy()
    dlv = sales_df[(sales_df['Статус'] == 'Доставлен') & (sales_df['Артикул'].isin(selected_arts))].copy()
    canc = sales_df[(sales_df['Статус'] == 'Отменён') & (sales_df['Артикул'].isin(selected_arts))].copy()

    # Выручка по артикулу
    rev_by_art = dlv.groupby('Артикул')['Оплачено покупателем'].sum().sort_values(ascending=False)
    total_by_art = dlv.groupby('Артикул')['Количество'].sum().sort_values(ascending=False)

    # Средний чек по артикулу
    avg_check_by_art = {}
    for art in selected_arts:
        art_dlv = dlv[dlv['Артикул'] == art]
        qty = art_dlv['Количество'].sum()
        if qty > 0:
            avg_check_by_art[art] = art_dlv['Оплачено покупателем'].sum() / qty

    # ── Лидер продаж ──
    if len(total_by_art) >= 2:
        leader = total_by_art.index[0]
        leader_qty = int(total_by_art.iloc[0])
        share = leader_qty / total_by_art.sum() * 100
        leader_rev = rev_by_art.get(leader, 0)
        second = total_by_art.index[1]
        second_qty = int(total_by_art.iloc[1])
        gap = leader_qty - second_qty
        insights.append(
            f"🏆 **Лидер продаж — {leader}**: {_fmt(leader_qty)} доставок ({share:.0f}%), "
            f"выручка {_fmt(leader_rev)} ₽. "
            f"Опережает {second} на {_fmt(gap)} шт."
        )
    elif not total_by_art.empty:
        leader = total_by_art.index[0]
        leader_qty = int(total_by_art.iloc[0])
        leader_rev = rev_by_art.get(leader, 0)
        insights.append(
            f"🏆 **Лидер продаж — {leader}**: {_fmt(leader_qty)} доставок, "
            f"выручка {_fmt(leader_rev)} ₽."
        )

    # ── Средний чек ──
    if avg_check_by_art:
        best_check_art = max(avg_check_by_art, key=avg_check_by_art.get)
        worst_check_art = min(avg_check_by_art, key=avg_check_by_art.get)
        if best_check_art != worst_check_art:
            insights.append(
                f"💰 **Средний чек**: самый высокий у {best_check_art} — "
                f"{avg_check_by_art[best_check_art]:,.0f} ₽/шт, ".replace(',', '\u202f') +
                f"самый низкий у {worst_check_art} — "
                f"{avg_check_by_art[worst_check_art]:,.0f} ₽/шт.".replace(',', '\u202f')
            )

    # ── Конверсия: % отмен ──
    canc_by_art = canc.groupby('Артикул')['Количество'].sum()
    orders_by_art = sales_df[sales_df['Артикул'].isin(selected_arts)].groupby('Артикул')['Количество'].sum()
    cancel_rate = {}
    for art in selected_arts:
        total_ord = orders_by_art.get(art, 0)
        total_canc = canc_by_art.get(art, 0)
        if total_ord > 0:
            cancel_rate[art] = total_canc / total_ord * 100
    if cancel_rate:
        worst_cancel_art = max(cancel_rate, key=cancel_rate.get)
        worst_rate = cancel_rate[worst_cancel_art]
        avg_rate = sum(cancel_rate.values()) / len(cancel_rate)
        if worst_rate > avg_rate * 1.3 and worst_rate > 10:
            insights.append(
                f"🔴 **{worst_cancel_art} — высокий % отмен**: {worst_rate:.0f}% заказов отменено "
                f"(средний по артикулам: {avg_rate:.0f}%). Возможная причина: описание, фото или сроки доставки."
            )

    # ── Стокаут-потери (по остатку на конец периода) ──
    # Средние продажи за нормальные месяцы (остаток >= продаж)
    for art in selected_arts:
        art_an = an[an['Артикул'] == art].sort_values('period_start')
        if len(art_an) < 2:
            continue
        normal_dlv = []
        for _, r in art_an.iterrows():
            if pd.notna(r['остаток_конец']) and pd.notna(r['Доставлено']) and r['Доставлено'] > 0:
                if r['остаток_конец'] >= r['Доставлено']:
                    normal_dlv.append(r['Доставлено'])
        avg_sales = sum(normal_dlv) / len(normal_dlv) if normal_dlv else 0
        if avg_sales <= 0:
            continue
        last_row = art_an.iloc[-1]
        last_stock = last_row['остаток_конец']
        last_dlv = last_row['Доставлено'] if pd.notna(last_row['Доставлено']) else 0
        if pd.notna(last_stock) and last_stock < avg_sales and last_dlv < avg_sales * 0.5:
            missed = int(avg_sales - last_dlv)
            avg_price = avg_check_by_art.get(art, 0)
            lost_rev = missed * avg_price
            lost_str = f" Упущенная выручка: ~{_fmt(lost_rev)} ₽/мес." if avg_price > 0 else ""
            insights.append(
                f"⚠️ **{art} — стокаут съел продажи**: "
                f"средние продажи {avg_sales:.0f} шт/мес, "
                f"в {last_row['month_label']} продано {_fmt(last_dlv)} шт (остаток {_fmt(last_stock)} шт). "
                f"Недопродано ~{_fmt(missed)} шт.{lost_str}"
            )

    # ── Восстановление после пополнения ──
    if an.empty or 'period_start' not in an.columns:
        return insights
    last_month = an['period_start'].max()
    prev_month = an[an['period_start'] < last_month]['period_start'].max()
    last_label = an[an['period_start'] == last_month]['month_label'].iloc[0] if len(an[an['period_start'] == last_month]) > 0 else ''
    prev_label = an[an['period_start'] == prev_month]['month_label'].iloc[0] if pd.notna(prev_month) and len(an[an['period_start'] == prev_month]) > 0 else ''
    if pd.notna(prev_month):
        last = an[an['period_start'] == last_month].groupby('Артикул')['Доставлено'].sum()
        prev = an[an['period_start'] == prev_month].groupby('Артикул')['Доставлено'].sum()
        common = last.index.intersection(prev.index)
        for art in common:
            prev_val = float(prev[art])
            last_val = float(last[art])
            if prev_val > 0 and last_val / prev_val >= 1.8:
                growth = (last_val / prev_val - 1) * 100
                abs_growth = int(last_val) - int(prev_val)
                insights.append(
                    f"📈 **{art} — резкий рост**: "
                    f"с {_fmt(prev_val)} ({prev_label}) до {_fmt(last_val)} ({last_label}) доставок "
                    f"(+{_fmt(abs_growth)} шт, +{growth:.0f}%)."
                )

    # ── Общий тренд ──
    last_an = an[an['period_start'] == last_month].drop_duplicates('Артикул') if pd.notna(last_month) else pd.DataFrame()
    avg_monthly = an.groupby('Артикул')['Доставлено'].mean()
    dlv['month_dt'] = dlv['date'].dt.to_period('M').apply(lambda p: p.start_time)
    monthly_total = dlv.groupby('month_dt')['Количество'].sum().sort_index()
    if len(monthly_total) >= 4:
        first_months = monthly_total.iloc[:3]
        last_months = monthly_total.iloc[-3:]
        first_avg = first_months.mean()
        last_avg = last_months.mean()
        trend_pct = (last_avg / first_avg - 1) * 100
        direction = "вырос" if trend_pct > 0 else "упал"
        emoji = "📊" if trend_pct > 0 else "📉"
        first_period = f"{first_months.index[0].strftime('%b %Y')}–{first_months.index[-1].strftime('%b %Y')}"
        last_period = f"{last_months.index[0].strftime('%b %Y')}–{last_months.index[-1].strftime('%b %Y')}"
        insights.append(
            f"{emoji} **Общий тренд {direction} на {abs(trend_pct):.0f}%**: "
            f"среднее {_fmt(first_avg)} шт/мес ({first_period}) → {_fmt(last_avg)} шт/мес ({last_period})."
        )

    # ── Нулевой остаток ──
    zero_stock = last_an[last_an['остаток_конец'] == 0]['Артикул'].tolist() if not last_an.empty and 'остаток_конец' in last_an.columns else []
    if zero_stock:
        zero_lost = sum(avg_monthly.get(a, 0) * avg_check_by_art.get(a, 0) for a in zero_stock)
        zero_lost_str = f" Потенциальные потери: ~{_fmt(zero_lost)} ₽/мес." if zero_lost > 0 else ""
        insights.append(
            f"🚨 **Нулевой остаток**: {', '.join(zero_stock)} — "
            f"продажи остановлены из-за отсутствия товара.{zero_lost_str}"
        )

    return insights


with st.expander('💡 Аналитические наблюдения', expanded=True):
    insights = build_insights(sales, analytics, selected)
    if insights:
        for note in insights:
            st.markdown(f"- {note}")
    else:
        st.markdown('_Нет данных для анализа._')

st.divider()

# ── График 1 + тепловая карта (общий subplot) ────────────────────────────────

title_col, info_col = st.columns([10, 1])
title_col.subheader('Динамика продаж')
with info_col.popover('ℹ️'):
    st.markdown("""
**Динамика продаж — Доставлено, шт**

Линии показывают количество доставленных товаров по каждому артикулу за период (неделя/месяц).

**Итого** (жирная линия) — суммарные продажи по всем выбранным артикулам.

Дельты в KPI показывают изменение за последний месяц по сравнению с предыдущим.
""")


# Агрегируем продажи
df_plot = delivered_sel.copy()
df_plot['period'] = (
    df_plot['date'].dt.to_period('W').apply(lambda p: p.start_time)
    if granularity == 'Неделя'
    else df_plot['date'].dt.to_period('M').apply(lambda p: p.start_time)
)
grouped = (
    df_plot.groupby(['period', 'Артикул'])['Количество']
    .sum()
    .reset_index()
)

analytics_sel = analytics[analytics['Артикул'].isin(selected)]
month_order = analytics_sel.drop_duplicates('month_label').sort_values('period_start')['month_label'].tolist()


COLORS = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
    '#aec7e8', '#ffbb78',
]
# Цвет привязан к артикулу глобально (по позиции в all_articles), а не к индексу в текущей
# выборке. Иначе один и тот же 06SK2024 в разных графиках получает разные цвета.
ARTICLE_COLOR = {art: COLORS[i % len(COLORS)] for i, art in enumerate(all_articles)}

# График продаж (без subplot — отдельный)
fig = go.Figure()

# Линия «Итого» — жирная, поверх остальных
selected_sorted = [a for a in all_articles if a in selected]
total_data = grouped[grouped['Артикул'].isin(selected)].groupby('period')['Количество'].sum().reset_index().sort_values('period')
if len(selected_sorted) > 1:
    fig.add_trace(
        go.Scatter(
            x=total_data['period'],
            y=total_data['Количество'],
            name='Итого',
            mode='lines+markers',
            line=dict(width=3.5, color='#333', dash='solid'),
            marker=dict(size=6, symbol='diamond'),
            opacity=0.85,
        ),
    )

# Линии продаж по артикулам (тоньше)
for i, art in enumerate(selected_sorted):
    art_data = grouped[grouped['Артикул'] == art].sort_values('period')
    fig.add_trace(
        go.Scatter(
            x=art_data['period'],
            y=art_data['Количество'],
            name=art,
            mode='lines+markers',
            line=dict(width=1.5, color=ARTICLE_COLOR[art]),
            marker=dict(size=4),
            opacity=0.7,
        ),
    )

fig.update_layout(
    height=420,
    hovermode='x unified',
    yaxis_title='Доставлено, шт',
    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0),
    margin=dict(t=40, b=20),
)

st.plotly_chart(fig, use_container_width=True)

# ── Упущенная выручка из-за стокаутов ──────────────────────────────────────

title_lost, info_lost = st.columns([10, 1])
title_lost.subheader('💸 Упущенная выручка из-за стокаутов')
with info_lost.popover('ℹ️'):
    st.markdown("""
**Как считается:**

Стокаут = месяц, где **остаток на конец < средних продаж** (товар заканчивался).

Для таких месяцев:

`упущено = (средние_продажи − фактические_продажи) × средний_чек`

**Средние продажи** считаются по «нормальным» месяцам — где остаток на конец ≥ продаж (товар не заканчивался, продажи не были ограничены стоком).

Пример: средние продажи = 100 шт, а в марте продали только 4 шт (остаток 1) → упущено 96 шт × средний чек.
""")

# Средний чек по артикулу
_avg_check = {}
for art in selected:
    art_dlv = delivered_sel[delivered_sel['Артикул'] == art]
    qty = art_dlv['Количество'].sum()
    if qty > 0:
        _avg_check[art] = art_dlv['Оплачено покупателем'].sum() / qty

# Pivot: остаток на конец месяца
pivot_stock_end = analytics_sel.pivot_table(
    index='Артикул', columns='month_label',
    values='остаток_конец', aggfunc='first'
).reindex(columns=month_order).reindex(selected_sorted)

# Pivot: доставки по месяцам
pivot_dlv = analytics_sel.pivot_table(
    index='Артикул', columns='month_label',
    values='Доставлено', aggfunc='first'
).reindex(columns=month_order).reindex(selected_sorted)

# Рассчитываем упущенную выручку
# Логика: если остаток_конец < средних_продаж → товар закончился раньше конца месяца
# Упущено = (средние_продажи − фактические_продажи) × средний_чек
lost_data = {}
for art in selected_sorted:
    if art not in pivot_stock_end.index or art not in pivot_dlv.index:
        continue

    stock_row = pivot_stock_end.loc[art]
    dlv_row = pivot_dlv.loc[art]
    avg_check = _avg_check.get(art, 0)

    # Шаг 1: определяем «нормальные» месяцы — где остаток_конец >= продаж
    # (товар не заканчивался, продажи не ограничены стоком)
    all_dlv = []
    for month in month_order:
        stock_end = stock_row.get(month, None)
        dlv_val = dlv_row.get(month, 0)
        if pd.notna(dlv_val) and dlv_val > 0:
            all_dlv.append(dlv_val)

    # Средние продажи — берём из месяцев где остаток >= доставлено (товар не заканчивался)
    normal_months_dlv = []
    for month in month_order:
        stock_end = stock_row.get(month, None)
        dlv_val = dlv_row.get(month, 0)
        if pd.isna(dlv_val) or dlv_val <= 0:
            continue
        if pd.notna(stock_end) and stock_end >= dlv_val:
            normal_months_dlv.append(dlv_val)

    avg_monthly_sales = sum(normal_months_dlv) / len(normal_months_dlv) if normal_months_dlv else 0

    # Шаг 2: считаем потери.
    # ВАЖНО: NaN в остаток_конец означает «товара физически нет на складе» (в xlsx стоит «–»),
    # а не «нет данных». Поэтому NaN-остаток обязан давать стокаут, иначе из расчётов
    # выпадают полные стокауты вроде 06SK2024 в Apr 2026 (история была — 0 продаж не из-за спроса).
    art_lost = {}
    for month in month_order:
        stock_end = stock_row.get(month, None)
        dlv_val = dlv_row.get(month, 0)
        if pd.isna(dlv_val):
            dlv_val = 0
        if avg_monthly_sales <= 0:
            # Нет истории нормальных продаж (новый/неактивный товар) — потерь нет
            art_lost[month] = 0
            continue
        # Остаток >= средних продаж → товар не заканчивался;
        # иначе (включая NaN — товара нет совсем) — стокаут.
        if pd.notna(stock_end) and stock_end >= avg_monthly_sales:
            art_lost[month] = 0
        else:
            missed_sales = max(avg_monthly_sales - dlv_val, 0)
            art_lost[month] = missed_sales * avg_check

    lost_data[art] = art_lost

# Столбиковый график — упущенная выручка по месяцам
fig_lost = go.Figure()
for i, art in enumerate(selected_sorted):
    if art not in lost_data:
        continue
    values = [lost_data[art].get(m, 0) for m in month_order]
    if sum(values) == 0:
        continue
    fig_lost.add_trace(go.Bar(
        name=art,
        x=month_order,
        y=values,
        marker_color=ARTICLE_COLOR[art],
        text=[f'{v:,.0f}' if v > 0 else '' for v in values],
        textposition='auto',
        textfont_size=10,
    ))

total_lost_all = sum(sum(v.values()) for v in lost_data.values())

# Аннотации — итого над каждым столбиком
month_totals = {}
for month in month_order:
    month_totals[month] = sum(lost_data.get(art, {}).get(month, 0) for art in selected_sorted)

annotations = []
for month in month_order:
    total = month_totals[month]
    if total > 0:
        annotations.append(dict(
            x=month, y=total,
            text=f'<b>{total:,.0f}</b>'.replace(',', ' '),
            showarrow=False,
            yshift=12,
            font=dict(size=11, color='#333'),
        ))

fig_lost.update_layout(
    barmode='stack',
    height=380,
    yaxis_title='Упущено, ₽',
    hovermode='x unified',
    legend=dict(orientation='h', yanchor='bottom', y=1.02),
    margin=dict(t=40, b=20),
    annotations=annotations,
)
st.plotly_chart(fig_lost, use_container_width=True)

# Итог + топ потерь
st.markdown(f'**Итого упущено за весь период: {total_lost_all:,.0f} ₽**'.replace(',', ' '))

top_lost = sorted(
    [(art, sum(months.values())) for art, months in lost_data.items()],
    key=lambda x: x[1], reverse=True
)
top3 = [f'**{art}** — {_fmt(total)} ₽' for art, total in top_lost[:3] if total > 0]
if top3:
    st.caption('Топ потерь: ' + ' · '.join(top3))

# ── График 2: Остатки на конец месяца + Приходы ────────────────────────────

st.subheader('Остаток на конец месяца и приходы')

pivot_stock = analytics_sel.pivot_table(
    index='Артикул', columns='month_label',
    values='остаток_конец', aggfunc='first'
).reindex(columns=month_order).reindex(selected_sorted)

pivot_dlv_stock = analytics_sel.pivot_table(
    index='Артикул', columns='month_label',
    values='Доставлено', aggfunc='first'
).reindex(columns=month_order).reindex(selected_sorted)

# Рассчитываем приходы: приход = остаток_конец - остаток_предыдущего_месяца + доставлено
# Порог: < 20 шт считаем шумом (возвраты, корректировки), а не реальной поставкой
INCOMING_THRESHOLD = 20
pivot_incoming = pd.DataFrame(index=pivot_stock.index, columns=month_order, dtype=float)
for art in selected_sorted:
    if art not in pivot_stock.index:
        continue
    for j, month in enumerate(month_order):
        stock_end = pivot_stock.loc[art].get(month, None)
        delivered = pivot_dlv_stock.loc[art].get(month, 0)
        if pd.isna(stock_end):
            pivot_incoming.loc[art, month] = None
            continue
        if pd.isna(delivered):
            delivered = 0
        if j == 0:
            pivot_incoming.loc[art, month] = None
        else:
            prev_month = month_order[j - 1]
            stock_prev = pivot_stock.loc[art].get(prev_month, None)
            if pd.isna(stock_prev):
                pivot_incoming.loc[art, month] = None
            else:
                incoming = stock_end - stock_prev + delivered
                # Отсекаем шум: мелкие возвраты/корректировки — не поставки
                pivot_incoming.loc[art, month] = incoming if incoming >= INCOMING_THRESHOLD else 0

fig2 = make_subplots(
    rows=2, cols=1,
    shared_xaxes=True,
    row_heights=[0.55, 0.45],
    vertical_spacing=0.08,
    subplot_titles=['Остаток на конец месяца', 'Приходы за месяц'],
)

for art in selected_sorted:
    if art not in pivot_stock.index:
        continue
    fig2.add_trace(go.Bar(
        name=art,
        x=month_order,
        y=pivot_stock.loc[art].values,
        marker_color=ARTICLE_COLOR[art],
        legendgroup=art,
    ), row=1, col=1)

    incoming_vals = pivot_incoming.loc[art].values if art in pivot_incoming.index else []
    fig2.add_trace(go.Bar(
        name=art,
        x=month_order,
        y=incoming_vals,
        marker_color=ARTICLE_COLOR[art],
        marker_line_width=1,
        marker_line_color='white',
        marker_opacity=0.7,
        legendgroup=art,
        showlegend=False,
    ), row=2, col=1)

fig2.update_layout(
    barmode='group',
    height=520,
    hovermode='x unified',
    legend=dict(orientation='h', yanchor='bottom', y=1.02),
    margin=dict(t=40, b=20),
)
fig2.update_yaxes(title_text='шт', row=1, col=1)
fig2.update_yaxes(title_text='шт', row=2, col=1)
st.plotly_chart(fig2, use_container_width=True)

# ── Анализ запасов (карточки) ────────────────────────────────────────────────

st.subheader('Анализ запасов')

last_month_an = analytics_sel[analytics_sel['period_start'] == analytics_sel['period_start'].max()]
avg_monthly_dlv = analytics_sel.groupby('Артикул')['Доставлено'].mean()

# Собираем данные без дубликатов
stock_cards = {}
for _, row in last_month_an.iterrows():
    art = row['Артикул']
    if art in stock_cards:
        continue
    stock = row['остаток_конец']
    avg = avg_monthly_dlv.get(art, 0)
    if avg <= 0:
        # У артикула нет истории продаж — карточку строить не из чего.
        continue
    # NaN в остаток_конец означает «товара физически нет на складе» (в xlsx стоит «–»),
    # а не «нет данных» — тот же инвариант, что в расчёте упущенной выручки.
    # Если пропустить такие строки — лидер продаж с пустым складом (06SK2024 в Apr 2026)
    # исчезает из «Анализа запасов», хотя это и есть самый срочный стокаут.
    if pd.isna(stock):
        stock = 0
    months_cover = stock / avg
    avg_price = delivered_sel[delivered_sel['Артикул'] == art]
    if len(avg_price) > 0:
        total_rev = avg_price['Оплачено покупателем'].sum()
        total_qty = avg_price['Количество'].sum()
        price_per = total_rev / total_qty if total_qty > 0 else 0
    else:
        price_per = 0
    frozen = stock * price_per
    frozen_cost = stock * COST_PER_UNIT
    stock_cards[art] = {
        'stock': int(stock),
        'avg': avg,
        'months': months_cover,
        'frozen': frozen,
        'frozen_cost': frozen_cost,
        'price': price_per,
    }

# Сортировка по срочности действий: стокаут → заканчивается → перезатарка → избыток → норма.
# Внутри группы — по абсолютной величине проблемы:
# - для стокаутов/заканчивающихся сортируем по упущенной выручке (худший наверху);
# - для перезатарки/избытка — по объёму замороженных денег;
# - для нормы — по выручке (просто чтобы был стабильный порядок).
def _card_sort_key(item):
    m = item[1]['months']
    if m < 0.5:
        priority = 0  # стокаут — действовать срочно, нужна поставка
    elif m < 1:
        priority = 1  # заканчивается — пора заказывать
    elif m > 6:
        priority = 2  # перезатарка — много заморожено
    elif m > 3:
        priority = 3  # избыток
    else:
        priority = 4  # норма
    return (priority, -item[1]['frozen'])

sorted_cards = sorted(stock_cards.items(), key=_card_sort_key)

# Сводка «Заморожено в товаре» — наверху раздела
total_frozen = sum(d['frozen'] for _, d in sorted_cards)
total_frozen_cost = sum(d['frozen_cost'] for _, d in sorted_cards)
total_diff = total_frozen - total_frozen_cost
total_stock = sum(d['stock'] for _, d in sorted_cards)

st.markdown(
    f'<div style="background:linear-gradient(135deg,#f8f9ff,#eef1ff);border:1px solid #d0d5f0;'
    f'border-radius:14px;padding:20px 28px;margin-bottom:20px">'
    f'<div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:16px">'
    f'<div style="text-align:center;flex:1;min-width:140px">'
    f'<div style="font-size:12px;color:#888;text-transform:uppercase;letter-spacing:1px">Всего на складе</div>'
    f'<div style="font-size:28px;font-weight:800;color:#333">{_fmt(total_stock)} шт</div>'
    f'</div>'
    f'<div style="text-align:center;flex:1;min-width:140px">'
    f'<div style="font-size:12px;color:#888;text-transform:uppercase;letter-spacing:1px">Розничная цена</div>'
    f'<div style="font-size:28px;font-weight:800;color:#e74c3c">{_fmt(total_frozen)} ₽</div>'
    f'</div>'
    f'<div style="text-align:center;flex:1;min-width:140px">'
    f'<div style="font-size:12px;color:#888;text-transform:uppercase;letter-spacing:1px">Себестоимость ({COST_PER_UNIT} ₽/шт)</div>'
    f'<div style="font-size:28px;font-weight:800;color:#2980b9">{_fmt(total_frozen_cost)} ₽</div>'
    f'</div>'
    f'<div style="text-align:center;flex:1;min-width:140px">'
    f'<div style="font-size:12px;color:#888;text-transform:uppercase;letter-spacing:1px">Наценка</div>'
    f'<div style="font-size:28px;font-weight:800;color:#27ae60">{_fmt(total_diff)} ₽</div>'
    f'</div>'
    f'</div>'
    f'</div>',
    unsafe_allow_html=True,
)

# Рендер карточек по 3 в ряд
for row_start in range(0, len(sorted_cards), 3):
    row_cards = sorted_cards[row_start:row_start + 3]
    cols = st.columns(3)
    for idx, (art, data) in enumerate(row_cards):
        with cols[idx]:
            months = data['months']
            # Нижние пороги критичнее перезатарки: пустой склад блокирует продажи,
            # а перезатарка — только замораживает деньги. Поэтому порядок: стокаут → норма → перезатарка.
            if months < 0.5:
                bar_color = '#c0392b'
                status = 'Стокаут'
                status_color = '#c0392b'
            elif months < 1:
                bar_color = '#e67e22'
                status = 'Заканчивается'
                status_color = '#e67e22'
            elif months <= 3:
                bar_color = '#2ecc71'
                status = 'Норма'
                status_color = '#2ecc71'
            elif months <= 6:
                bar_color = '#f39c12'
                status = 'Избыток'
                status_color = '#e67e22'
            else:
                bar_color = '#e74c3c'
                status = 'Перезатарка'
                status_color = '#e74c3c'

            bar_pct = min(months / 15 * 100, 100)
            img = ARTICLE_IMAGES.get(art, '')

            st.markdown(
                f'<div style="border:1px solid #e0e0e0;border-radius:12px;padding:16px;margin-bottom:8px">'
                f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:12px">'
                f'<img src="{img}" style="width:56px;height:56px;border-radius:8px;object-fit:cover" onerror="this.style.display=\'none\'">'
                f'<div>'
                f'<div style="font-weight:700;font-size:16px">{art}</div>'
                f'<span style="background:{status_color};color:white;padding:2px 8px;border-radius:10px;font-size:12px">{status}</span>'
                f'</div>'
                f'</div>'
                f'<div style="display:flex;justify-content:space-between;font-size:13px;color:#666;margin-bottom:4px">'
                f'<span>Остаток: <b>{_fmt(data["stock"])} шт</b></span>'
                f'<span>Темп: <b>{data["avg"]:.0f} шт/мес</b></span>'
                f'</div>'
                f'<div style="background:#f0f0f0;border-radius:6px;height:14px;margin:8px 0;overflow:hidden">'
                f'<div style="background:{bar_color};height:100%;width:{bar_pct:.0f}%;border-radius:6px"></div>'
                f'</div>'
                f'<div style="display:flex;justify-content:space-between;font-size:13px">'
                f'<span style="color:{status_color};font-weight:600">Запас: {months:.1f} мес</span>'
                f'<span style="color:#888;text-align:right;font-size:12px">розн. ~{_fmt(data["frozen"])} ₽<br>с/с ~{_fmt(data["frozen_cost"])} ₽</span>'
                f'</div>'
                f'</div>',
                unsafe_allow_html=True,
            )


st.divider()

# ── % отмен по артикулам ────────────────────────────────────────────────────

st.subheader('📊 Процент отмен по артикулам')

_orders_by_art = sales[sales['Артикул'].isin(selected)].groupby('Артикул')['Количество'].sum()
_canc_by_art = cancelled_sel.groupby('Артикул')['Количество'].sum()
cancel_data = []
for art in selected_sorted:
    total_ord = _orders_by_art.get(art, 0)
    total_canc = _canc_by_art.get(art, 0)
    rate = (total_canc / total_ord * 100) if total_ord > 0 else 0
    cancel_data.append({'art': art, 'rate': rate, 'cancelled': total_canc, 'total': total_ord})

cancel_data.sort(key=lambda x: x['rate'], reverse=True)

fig_canc = go.Figure()
colors_canc = ['#e74c3c' if d['rate'] > 15 else '#f39c12' if d['rate'] > 10 else '#2ecc71' for d in cancel_data]
fig_canc.add_trace(go.Bar(
    x=[d['art'] for d in cancel_data],
    y=[d['rate'] for d in cancel_data],
    marker_color=colors_canc,
    text=[f"{d['rate']:.1f}%<br>({int(d['cancelled'])} шт)" for d in cancel_data],
    textposition='outside',
    textfont_size=11,
))
avg_cancel = sum(d['rate'] for d in cancel_data) / len(cancel_data) if cancel_data else 0
fig_canc.add_hline(y=avg_cancel, line_dash='dash', line_color='#888',
                   annotation_text=f'Среднее: {avg_cancel:.1f}%', annotation_position='top left')
fig_canc.update_layout(
    height=300,
    yaxis_title='% отмен',
    margin=dict(t=20, b=20),
    showlegend=False,
)
st.plotly_chart(fig_canc, use_container_width=True)

st.divider()

# ── Сводная таблица ──────────────────────────────────────────────────────────

with st.expander('Сводная таблица по месяцам'):
    tbl = (
        analytics_sel[['month_label', 'Артикул', 'Заказано', 'Доставлено',
                        'Отменено', 'остаток_конец']]
        .rename(columns={
            'month_label': 'Месяц',
            'остаток_конец': 'Остаток',
        })
        .sort_values(['Месяц', 'Артикул'])
        .reset_index(drop=True)
    )
    # Форматирование с цветовой индикацией
    num_cols = ['Заказано', 'Доставлено', 'Отменено', 'Остаток']
    for c in num_cols:
        tbl[c] = pd.to_numeric(tbl[c], errors='coerce')

    def _highlight_stock(val):
        if pd.isna(val):
            return ''
        if val == 0:
            return 'background-color: #ffcccc; color: #c0392b; font-weight: 700'
        if val < 10:
            return 'background-color: #fff3cd; color: #856404'
        return ''

    styled = (
        tbl.style
        .format({c: '{:,.0f}' for c in num_cols}, na_rep='—')
        .map(_highlight_stock, subset=['Остаток'])
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)
