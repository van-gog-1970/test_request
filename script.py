import sys
from datetime import datetime

print("Запуск скрипта збереження тексту...")

# Перевіряємо, чи передав GitHub Actions нам текст
if len(sys.argv) > 1:
    user_text = sys.argv[1]
else:
    user_text = "Текст не було введено"

try:
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Режим "a" (append) дописує нові відповіді в кінець файлу, не видаляючи старі
    with open("user_answers.txt", "a", encoding="utf-8") as f:
        f.write(f"[{current_time}] Користувач ввів: {user_text}\n")
        
    print(f"Успішно збережено текст: '{user_text}' у файл user_answers.txt")
except Exception as e:
    print(f"Помилка при записі: {e}")
