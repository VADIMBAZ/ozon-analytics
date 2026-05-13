"""Загрузка файлов отчётов через UI дашборда и автокоммит в публичный mirror.

Поток:
    UI (st.file_uploader) → validate_upload() → commit_to_github()

Файлы коммитятся в `VADIMBAZ/ozon-analytics` через GitHub API (PyGithub),
а не через `git push` — Streamlit Cloud работает в stateless-контейнере,
локальной рабочей копии для CLI-команды нет. После коммита Streamlit Cloud
сам подхватит изменения и пересоберёт дашборд (~90 сек).

Перезаписанные файлы не плодят `.bak` рядом — git автоматически хранит
историю предыдущих версий (`git log -- <filename>`).
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Literal

import pandas as pd

FileKind = Literal['analytics', 'sales']

# Канонические русские названия месяцев (как в существующих xlsx).
RU_MONTHS = [
    '', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь',
]


class ValidationError(Exception):
    """Файл невалидный — наружу с понятным человеку сообщением."""


@dataclass
class ValidatedFile:
    kind: FileKind
    canonical_name: str   # как файл должен называться в репо
    content: bytes        # сырые байты для коммита
    period_label: str     # «Май 2026» / «01.05.2026–31.05.2026» — для пользователя


# ─────────────────────────────────────────────────────────────────────────
# Валидация и нормализация имён
# ─────────────────────────────────────────────────────────────────────────


def _validate_analytics_xlsx(name: str, raw: bytes) -> ValidatedFile:
    """Проверяем, что xlsx — действительно analytics_report от Ozon, и определяем период."""
    try:
        df = pd.read_excel(io.BytesIO(raw), header=None)
    except Exception as e:
        raise ValidationError(f'Не удалось прочитать xlsx: {e}')

    if df.shape[0] < 14 or df.shape[1] < 29:
        raise ValidationError(
            'xlsx не похож на analytics_report: ожидаются 14+ строк и 29+ колонок, '
            f'а в файле {df.shape[0]}×{df.shape[1]}.'
        )

    period_str = str(df.iloc[0, 0])
    dates = re.findall(r'(\d{2}\.\d{2}\.\d{4})', period_str)
    if len(dates) < 2:
        raise ValidationError(
            f'В заголовке xlsx (ячейка A1) не нашёл двух дат периода. '
            f'Содержимое: «{period_str[:120]}». Похоже, это не analytics_report.'
        )

    period_start = datetime.strptime(dates[0], '%d.%m.%Y')
    period_end = datetime.strptime(dates[1], '%d.%m.%Y')

    # Проверка, что период — ровно один месяц (Ozon выдаёт помесячные отчёты).
    if period_start.day != 1 or period_end.month != period_start.month or period_end.year != period_start.year:
        raise ValidationError(
            f'Период в файле — с {dates[0]} по {dates[1]}, это не один календарный месяц. '
            'Ozon-аналитика должна выгружаться помесячно (с 1-го по последний день месяца).'
        )

    month_label = f'{RU_MONTHS[period_start.month]} {period_start.year}'
    canonical_name = f'analytics_report_{month_label}.xlsx'

    return ValidatedFile(
        kind='analytics',
        canonical_name=canonical_name,
        content=raw,
        period_label=month_label,
    )


def _validate_sales_file(name: str, raw: bytes) -> ValidatedFile:
    """Проверяем, что файл заказов — корректный (по составу колонок)."""
    ext = name.lower().rsplit('.', 1)[-1] if '.' in name else ''
    if ext not in ('numbers', 'csv', 'xlsx'):
        raise ValidationError(f'Неподдерживаемое расширение «.{ext}». Ожидаются: numbers/csv/xlsx.')

    # Пытаемся распарсить, чтобы извлечь даты и проверить колонки.
    df = _parse_sales_bytes(raw, ext)

    required_cols = {'Принят в обработку', 'Артикул', 'Количество', 'Оплачено покупателем', 'Статус'}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValidationError(
            'Файл заказов не содержит обязательных колонок: '
            f'{sorted(missing)}. Это похоже на не тот отчёт.'
        )

    dates = pd.to_datetime(df['Принят в обработку'], errors='coerce').dropna()
    if dates.empty:
        raise ValidationError('В колонке «Принят в обработку» нет распознаваемых дат.')

    date_from, date_to = dates.min(), dates.max()
    period_label = f"{date_from:%d.%m.%Y}–{date_to:%d.%m.%Y}"
    canonical_name = f'sales results {date_from:%d.%m.%Y}-{date_to:%d.%m.%Y}.{ext}'

    return ValidatedFile(
        kind='sales',
        canonical_name=canonical_name,
        content=raw,
        period_label=period_label,
    )


def _parse_sales_bytes(raw: bytes, ext: str) -> pd.DataFrame:
    """Парсит байты файла заказов. Любые ошибки парсинга → ValidationError."""
    try:
        if ext == 'csv':
            return pd.read_csv(io.BytesIO(raw), sep=';')
        if ext == 'xlsx':
            return pd.read_excel(io.BytesIO(raw))
        if ext == 'numbers':
            # numbers-parser принимает только пути → сохраняем во временный файл.
            import tempfile
            import numbers_parser
            with tempfile.NamedTemporaryFile(suffix='.numbers', delete=True) as tmp:
                tmp.write(raw)
                tmp.flush()
                doc = numbers_parser.Document(tmp.name)
                rows = list(doc.sheets[0].tables[0].rows())
                headers = [str(c.value) for c in rows[0]]
                data = [[c.value for c in row] for row in rows[1:]]
                return pd.DataFrame(data, columns=headers)
    except ValidationError:
        raise
    except Exception as e:
        raise ValidationError(f'Не удалось распарсить .{ext}: {e}')
    raise ValidationError(f'Парсер для .{ext} не реализован.')


def validate_upload(name: str, raw: bytes) -> ValidatedFile:
    """Определяет тип файла по расширению и валидирует.

    Args:
        name: имя файла, как пришло из st.file_uploader.
        raw: содержимое файла (bytes).

    Returns:
        ValidatedFile с каноническим именем и нормализованными данными.

    Raises:
        ValidationError: если файл не похож на ожидаемый отчёт.
    """
    lower = name.lower()
    if lower.endswith('.xlsx') and 'analytics' in lower:
        return _validate_analytics_xlsx(name, raw)
    if lower.endswith('.xlsx'):
        # xlsx без префикса analytics — может быть и sales (теоретически). Пробуем оба.
        try:
            return _validate_analytics_xlsx(name, raw)
        except ValidationError:
            return _validate_sales_file(name, raw)
    if lower.endswith(('.numbers', '.csv')):
        return _validate_sales_file(name, raw)
    raise ValidationError(
        f'Не понял тип файла «{name}». Ожидаются analytics_report_*.xlsx '
        'или sales results *.{numbers,csv,xlsx}.'
    )


# ─────────────────────────────────────────────────────────────────────────
# Commit в GitHub через PyGithub
# ─────────────────────────────────────────────────────────────────────────


def commit_to_github(
    file: ValidatedFile,
    *,
    repo_full_name: str,
    token: str,
    branch: str = 'main',
) -> dict:
    """Коммитит файл через GitHub Contents API. Возвращает информацию о коммите.

    При существующем файле — update (старая версия остаётся в git history).
    При новом файле — create.

    Returns:
        {'sha': '<commit-sha>', 'url': '<html_url>', 'action': 'created' | 'updated'}
    """
    from github import Github, GithubException

    gh = Github(token)
    repo = gh.get_repo(repo_full_name)
    path = file.canonical_name

    try:
        existing = repo.get_contents(path, ref=branch)
        message = (
            f'data: обновлён {file.period_label} '
            f'({"analytics" if file.kind == "analytics" else "sales"}) через дашборд'
        )
        result = repo.update_file(
            path=path,
            message=message,
            content=file.content,
            sha=existing.sha,
            branch=branch,
        )
        return {
            'action': 'updated',
            'sha': result['commit'].sha,
            'url': result['commit'].html_url,
        }
    except GithubException as e:
        if e.status != 404:
            raise
        # Файла нет — создаём.
        message = (
            f'data: добавлен {file.period_label} '
            f'({"analytics" if file.kind == "analytics" else "sales"}) через дашборд'
        )
        result = repo.create_file(
            path=path,
            message=message,
            content=file.content,
            branch=branch,
        )
        return {
            'action': 'created',
            'sha': result['commit'].sha,
            'url': result['commit'].html_url,
        }
