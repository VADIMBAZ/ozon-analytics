#!/bin/bash
cd "$(dirname "$0")"

if [ ! -d .venv ]; then
    echo "Создаю виртуальное окружение..."
    python3 -m venv .venv
fi

if [ ! -f .venv/bin/streamlit ]; then
    echo "Устанавливаю зависимости..."
    .venv/bin/pip install -r requirements.txt
fi

PORT="${DASH_PORT:-8501}"

case "${1:-}" in
    start)
        shift
        # Порт из аргумента или переменной
        if [ -n "${1:-}" ]; then PORT="$1"; fi
        pkill -f "streamlit run app/main.py" 2>/dev/null
        sleep 1
        nohup .venv/bin/streamlit run app/main.py --server.headless true --server.port "$PORT" > /tmp/streamlit.log 2>&1 &
        echo "Дашборд запущен в фоне (PID: $!)"
        echo "Открой: http://127.0.0.1:$PORT"
        ;;
    stop)
        pkill -f "streamlit run app/main.py" && echo "Дашборд остановлен" || echo "Дашборд не запущен"
        ;;
    status)
        if pgrep -f "streamlit run app/main.py" > /dev/null; then
            echo "Дашборд работает (PID: $(pgrep -f 'streamlit run app/main.py'))"
        else
            echo "Дашборд не запущен"
        fi
        ;;
    *)
        # Обычный запуск (с терминалом)
        echo "Запускаю дашборд..."
        echo "Совет: ./run.sh start — запуск в фоне, ./run.sh stop — остановка"
        .venv/bin/streamlit run app/main.py "$@"
        ;;
esac
