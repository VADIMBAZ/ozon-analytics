import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from services.loader import load_sales, load_analytics

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

st.title('📦 Спутник Ключи — Аналитика Ozon')

# ── KPI ─────────────────────────────────────────────────────────────────────

delivered_all = sales[sales['Статус'] == 'Доставлен']
cancelled_all = sales[sales['Статус'] == 'Отменён']

delivered_sel = delivered_all[delivered_all['Артикул'].isin(selected)]
cancelled_sel = cancelled_all[cancelled_all['Артикул'].isin(selected)]
total_orders = sales[sales['Артикул'].isin(selected)]['Количество'].sum()

col1, col2, col3, col4 = st.columns(4)
col1.metric('Всего заказано', f"{int(total_orders):,}".replace(',', ' '))
col2.metric('Доставлено', f"{int(delivered_sel['Количество'].sum()):,}".replace(',', ' '))
col3.metric('Отменено', f"{int(cancelled_sel['Количество'].sum()):,}".replace(',', ' '))
revenue = delivered_sel['Оплачено покупателем'].sum()
col4.metric('Выручка (доставл.)', f"{revenue:,.0f} ₽".replace(',', ' '))

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

    # ── Стокаут-потери ──
    avg_days = an.groupby('Артикул')['дней_без_остатка'].mean()
    stockout_arts = avg_days[avg_days > 15].sort_values(ascending=False)
    if not stockout_arts.empty:
        for worst in stockout_arts.index[:2]:
            worst_days = stockout_arts[worst]
            art_monthly = an[an['Артикул'] == worst].sort_values('period_start')
            if len(art_monthly) >= 2:
                peak_row = art_monthly.loc[art_monthly['Доставлено'].idxmax()]
                last_row = art_monthly.iloc[-1]
                drop = int(peak_row['Доставлено']) - int(last_row['Доставлено'])
                if drop > 0:
                    avg_price = avg_check_by_art.get(worst, 0)
                    lost_rev = drop * avg_price
                    lost_str = f" Упущенная выручка: ~{_fmt(lost_rev)} ₽/мес." if avg_price > 0 else ""
                    insights.append(
                        f"⚠️ **{worst} — стокаут съел продажи**: "
                        f"пик {_fmt(peak_row['Доставлено'])} шт в {peak_row['month_label']}, "
                        f"сейчас {_fmt(last_row['Доставлено'])} шт в {last_row['month_label']} "
                        f"(−{_fmt(drop)} шт, −{drop / int(peak_row['Доставлено']) * 100:.0f}%). "
                        f"Без остатка в среднем {worst_days:.0f} из 28 дней.{lost_str}"
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
**Верхний график — Доставлено, шт**

Линии показывают количество доставленных товаров по каждому артикулу за период (неделя/месяц).

🔴 **Красные зоны** на фоне — периоды массового стокаута (в среднем >15 дней без остатка). Падение линии в красной зоне = потерянные продажи из-за отсутствия товара.

---

**Нижний график — Дней без остатка из 28**

Тепловая карта: каждая ячейка — сколько дней в месяце артикул был **недоступен** (0 остаток на складе).

| Цвет | Значение |
|------|----------|
| 🟢 Зелёный | 0–5 дней — норма |
| 🟡 Жёлтый | 6–15 дней — проблемы с поставкой |
| 🔴 Красный | 16–28 дней — критический стокаут |

**Как читать:** если ячейка красная с числом 25 — товар отсутствовал 25 из 28 дней, продажи почти остановились.
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

# Месячные данные для тепловой карты
analytics_sel = analytics[analytics['Артикул'].isin(selected)]
month_order = analytics_sel.drop_duplicates('month_label').sort_values('period_start')['month_label'].tolist()
pivot_heat = analytics_sel.pivot_table(
    index='Артикул', columns='month_label',
    values='дней_без_остатка', aggfunc='first'
).reindex(columns=month_order).reindex([a for a in all_articles if a in selected])

# Зоны стокаутов для фона графика 1 (усреднённые по выбранным артикулам)
stockout_zones = (
    analytics_sel.groupby(['period_start', 'period_end'])['дней_без_остатка']
    .mean()
    .reset_index()
)

COLORS = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
    '#aec7e8', '#ffbb78',
]

# График продаж (без subplot — отдельный)
fig = go.Figure()

# Фон-зоны стокаутов
for _, z in stockout_zones.iterrows():
    if z['дней_без_остатка'] >= 15:
        opacity = min(0.25, z['дней_без_остатка'] / 28 * 0.35)
        fig.add_vrect(
            x0=z['period_start'], x1=z['period_end'],
            fillcolor=f'rgba(220,50,50,{opacity:.2f})',
            layer='below', line_width=0,
        )

# Линии продаж (порядок по продажам)
selected_sorted = [a for a in all_articles if a in selected]
for i, art in enumerate(selected_sorted):
    art_data = grouped[grouped['Артикул'] == art].sort_values('period')
    fig.add_trace(
        go.Scatter(
            x=art_data['period'],
            y=art_data['Количество'],
            name=art,
            mode='lines+markers',
            line=dict(width=2, color=COLORS[i % len(COLORS)]),
            marker=dict(size=5),
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

Стокаут = месяц, где **остаток на конец периода = 0** (товар закончился).

Для таких месяцев:

`упущено = средние_продажи_за_нормальные_месяцы × средний_чек`

«Нормальные месяцы» — те, где остаток на конец > 0 (товар был в наличии).

Поле `дней_без_остатка` из выгрузки OZON **не используется** — оно часто содержит ошибки.
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
).reindex(columns=month_order).reindex([a for a in all_articles if a in selected])

# Pivot: доставки по месяцам
pivot_dlv = analytics_sel.pivot_table(
    index='Артикул', columns='month_label',
    values='Доставлено', aggfunc='first'
).reindex(columns=month_order).reindex([a for a in all_articles if a in selected])

# Рассчитываем упущенную выручку
lost_data = {}
for art in selected_sorted:
    if art not in pivot_stock_end.index or art not in pivot_dlv.index:
        continue

    stock_row = pivot_stock_end.loc[art]
    dlv_row = pivot_dlv.loc[art]
    avg_check = _avg_check.get(art, 0)

    # Средние продажи за «нормальные» месяцы (остаток > 0)
    normal_months_dlv = []
    for month in month_order:
        stock_end = stock_row.get(month, None)
        dlv_val = dlv_row.get(month, 0)
        if pd.notna(stock_end) and stock_end > 0 and pd.notna(dlv_val) and dlv_val > 0:
            normal_months_dlv.append(dlv_val)

    avg_monthly_sales = sum(normal_months_dlv) / len(normal_months_dlv) if normal_months_dlv else 0

    # Считаем потери за месяцы со стокаутом
    art_lost = {}
    for month in month_order:
        stock_end = stock_row.get(month, None)
        if pd.isna(stock_end) or stock_end > 0:
            art_lost[month] = 0  # товар был — потерь нет
        else:
            # остаток = 0 → стокаут, считаем упущенное
            art_lost[month] = avg_monthly_sales * avg_check

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
        marker_color=COLORS[i % len(COLORS)],
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

# ── График 2: Остатки на конец месяца ────────────────────────────────────────

st.subheader('Остаток на конец месяца')

pivot_stock = analytics_sel.pivot_table(
    index='Артикул', columns='month_label',
    values='остаток_конец', aggfunc='first'
).reindex(columns=month_order).reindex(all_articles)

fig2 = go.Figure()
for i, art in enumerate(all_articles):
    if art not in pivot_stock.index:
        continue
    fig2.add_trace(go.Bar(
        name=art,
        x=month_order,
        y=pivot_stock.loc[art].values,
        marker_color=COLORS[i % len(COLORS)],
    ))

fig2.update_layout(
    barmode='group',
    height=320,
    yaxis_title='шт',
    hovermode='x unified',
    legend=dict(orientation='h', yanchor='bottom', y=1.02),
    margin=dict(t=40, b=20),
)
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
    if pd.isna(stock) or avg <= 0:
        continue
    months_cover = stock / avg
    avg_price = delivered_sel[delivered_sel['Артикул'] == art]
    if len(avg_price) > 0:
        total_rev = avg_price['Оплачено покупателем'].sum()
        total_qty = avg_price['Количество'].sum()
        price_per = total_rev / total_qty if total_qty > 0 else 0
    else:
        price_per = 0
    COST_PER_UNIT = 25
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

# Сортировка по замороженным деньгам
sorted_cards = sorted(stock_cards.items(), key=lambda x: x[1]['frozen'], reverse=True)

# Рендер карточек по 3 в ряд
for row_start in range(0, len(sorted_cards), 3):
    row_cards = sorted_cards[row_start:row_start + 3]
    cols = st.columns(3)
    for idx, (art, data) in enumerate(row_cards):
        with cols[idx]:
            months = data['months']
            if months <= 3:
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

total_frozen = sum(d['frozen'] for _, d in sorted_cards)
total_frozen_cost = sum(d['frozen_cost'] for _, d in sorted_cards)
total_diff = total_frozen - total_frozen_cost
st.markdown(
    f'**Итого заморожено в товаре:**\n\n'
    f'- Розничная цена: **{_fmt(total_frozen)} ₽**\n'
    f'- Себестоимость (25 ₽/шт): **{_fmt(total_frozen_cost)} ₽**\n'
    f'- Разница (наценка): **{_fmt(total_diff)} ₽**'
)

st.divider()

# ── Сводная таблица ──────────────────────────────────────────────────────────

with st.expander('Сводная таблица по месяцам'):
    tbl = (
        analytics_sel[['month_label', 'Артикул', 'Заказано', 'Доставлено',
                        'Отменено', 'дней_без_остатка', 'остаток_конец']]
        .rename(columns={
            'month_label': 'Месяц',
            'дней_без_остатка': 'Дней без остатка',
            'остаток_конец': 'Остаток конец',
        })
        .sort_values(['Месяц', 'Артикул'])
        .reset_index(drop=True)
    )
    st.dataframe(tbl, use_container_width=True, hide_index=True)
