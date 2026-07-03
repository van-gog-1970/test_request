mport urllib.request

print("Привіт! Скрипт успішно запустився на GitHub Actions.")

try:
    with urllib.request.urlopen("https://ifconfig.me") as response:
        ip = response.read().decode("utf-8")
        print(f"Поточна IP-адреса віртуальної машини: {ip}")
except Exception as e:
    print(f"Не вдалося визначити IP: {e}")
