from flask import Flask, render_template, jsonify
import pandas as pd
import random
import os

app = Flask(__name__)

# Загрузка данных из Excel
def load_events():
    """Загружает события из Excel файла"""
    try:
        # Читаем Excel файл
        df = pd.read_excel('BaseDNDDesert.xlsx', engine='openpyxl')
        
        # Преобразуем в список словарей
        events = df.to_dict('records')
        
        # Фильтруем пустые строки
        events = [event for event in events if pd.notna(event.get('Название'))]
        
        return events
    except Exception as e:
        print(f"Ошибка загрузки Excel: {e}")
        return []

# Загружаем события при старте
EVENTS = load_events()
print(f"Загружено {len(EVENTS)} событий")

@app.route('/')
def index():
    """Главная страница"""
    return render_template('index.html')

@app.route('/api/random-event')
def random_event():
    """Возвращает случайное событие"""
    if not EVENTS:
        return jsonify({'error': 'Нет событий в базе данных'}), 404
    
    event = random.choice(EVENTS)
    
    response = {
        'id': int(event.get('ID', 0)) if pd.notna(event.get('ID')) else random.randint(1, 1000),
        'title': str(event.get('Название', 'Без названия')),
        'category': str(event.get('Категория', 'Неизвестно')),
        'difficulty': str(event.get('Сложность', 'Неизвестно')),
        'danger': int(event.get('Опасность (1–10)', 3)) if pd.notna(event.get('Опасность (1–10)')) else 3,
        'reward': str(event.get('Награда', 'Нет награды')),
        'story': str(event.get('Сюжет', 'Описание отсутствует')),
        'development': str(event.get('Дальнейшее развитие', 'Нет информации')),
        'tags': str(event.get('Теги', ''))
    }
    
    return jsonify(response)

@app.route('/api/events-count')
def events_count():
    """Возвращает количество событий"""
    return jsonify({'count': len(EVENTS)})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)