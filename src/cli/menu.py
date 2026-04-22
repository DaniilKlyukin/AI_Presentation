import os
from pathlib import Path
from src.core.extractor import PptxToPpc
from src.core.modifier import CompactPptxModifier
from src.core.cleaner import clean_pptx_metadata
from src.core.creator import PptxCreator


def clean_path(path_str):
    return path_str.strip('" ').strip("' ")


def run_extraction():
    pptx_file = clean_path(input("Введите путь к исходному PPTX: "))
    if os.path.exists(pptx_file):
        out_file = pptx_file.replace('.pptx', '.txt')
        print("⏳ Извлечение данных...")
        extractor = PptxToPpc(pptx_file)
        extractor.extract(out_file)
        print(f"✅ Данные успешно извлечены в: {out_file}")
    else:
        print("❌ Файл не найден.")


def run_modification():
    pptx_in = clean_path(input("Путь к исходному PPTX: "))
    path_obj = Path(pptx_in)

    ppc_path_input = clean_path(input("Путь к TXT (Enter если файл рядом с PPTX): "))
    ppc_path = Path(ppc_path_input) if ppc_path_input else path_obj.with_suffix('.txt')

    out_path = path_obj.parent / f"modified_{path_obj.name}"

    if path_obj.exists() and ppc_path.exists():
        print("⏳ Применение изменений...")
        try:
            modifier = CompactPptxModifier(str(path_obj))
            slides_count = modifier.apply_ppc(str(ppc_path), str(out_path))
            print(f"🔍 Обраработано слайдов в файле: {slides_count}")
            print(f"✅ Готово! Результат сохранен в: {out_path}")
        except Exception as e:
            print(f"❌ Произошла ошибка при модификации: {e}")
    else:
        print("❌ Исходный PPTX или TXT файл не найден.")


def run_metadata_cleaner():
    target_dir = clean_path(input("Введите путь к папке (для рекурсивной очистки .pptx): "))

    if os.path.isdir(target_dir):
        confirm = input(f"Внимание! Метаданные будут перезаписаны в '{target_dir}'. Продолжить? (y/n): ")
        if confirm.lower() == 'y':
            print(f"⏳ Начинаю поиск файлов в: {target_dir}")
            processed, errors, details = clean_pptx_metadata(target_dir)

            for detail in details:
                print(detail)

            print("\n" + "=" * 30)
            print("Завершено!")
            print(f"Успешно обработано: {processed}")
            print(f"Ошибок: {errors}")
            print("=" * 30)
        else:
            print("Отмена операции.")
    else:
        print("❌ Указанный путь не является папкой или не существует.")


def run_creation():
    md_in = clean_path(input("Путь к TXT/MD файлу с текстом: "))
    path_obj = Path(md_in)

    if path_obj.exists():
        out_path = path_obj.parent / f"{path_obj.stem}.pptx"
        print("⏳ Создание презентации...")
        try:
            creator = PptxCreator()
            result = creator.create_from_file(str(path_obj), str(out_path))
            for warn in result["warnings"]:
                print(f"⚠️ ПРЕДУПРЕЖДЕНИЕ: {warn}")
            print(f"✅ Готово! Создано слайдов: {result['slides_created']}. Файл: {out_path}")
        except Exception as e:
            print(f"❌ Ошибка: {e}")
    else:
        print("❌ Файл не найден.")


def show_main_menu():
    while True:
        print("\n" + "=" * 40)
        print("PPTX Tool Suite - Главное меню")
        print("=" * 40)
        print("1. Создать новую презентацию из текста (TXT/MD)")  # НОВОЕ
        print("2. Извлечь данные из PPTX в TXT")
        print("3. Применить сложные изменения из TXT в PPTX")
        print("4. Очистить метаданные PPTX в папке")
        print("0. Выход")
        print("=" * 40)

        choice = input("Выберите действие: ").strip()

        if choice == '1':
            run_creation()
        elif choice == '2':
            run_extraction()
        elif choice == '3':
            run_modification()
        elif choice == '4':
            run_metadata_cleaner()
        elif choice == '0':
            print("Выход из программы.")
            break
        else:
            print("❌ Неверный ввод.")