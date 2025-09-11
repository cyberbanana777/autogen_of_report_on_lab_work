import os
from pathlib import Path

def list_images_in_folder(folder_path):
    """
    Выводит список всех изображений в указанной папке
    """
    # Поддерживаемые форматы изображений
    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'}
    
    # Получаем путь к текущей папке, если путь не указан
    if not folder_path:
        folder_path = os.getcwd()
    
    # Проверяем существование папки
    if not os.path.exists(folder_path):
        print(f"Папка '{folder_path}' не существует!")
        return
    
    print(f"Изображения в папке: {folder_path}")
    print("-" * 50)
    
    # Ищем файлы с изображениями
    image_files = []
    for file in os.listdir(folder_path):
        file_path = Path(file)
        if file_path.suffix.lower() in image_extensions:
            image_files.append(file)
    
    # Выводим результат
    if image_files:
        for i, image in enumerate(image_files, 1):
            print(f"{image}")
        print(f"\nНайдено изображений: {len(image_files)}")
    else:
        print("Изображения не найдены!")

# Использование
if __name__ == "__main__":
    # Укажите путь к папке или оставьте пустым для текущей папки
    folder_path = ""  # или "C:/путь/к/вашей/папке"
    list_images_in_folder(folder_path)