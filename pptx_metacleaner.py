import os
from pptx import Presentation
from pptx.exc import PythonPptxError


def clean_pptx_metadata(folder_path):
    """
    Рекурсивно обходит папку и очищает метаданные во всех pptx файлах.
    """
    files_processed = 0
    files_error = 0

    print(f"Начинаю поиск файлов в: {folder_path}")

    # os.walk позволяет заходить во все вложенные папки
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".pptx") and not file.startswith("~$"):
                file_path = os.path.join(root, file)

                try:
                    prs = Presentation(file_path)
                    props = prs.core_properties

                    # Очистка основных полей
                    props.title = ""
                    props.author = ""

                    # Дополнительная очистка (опционально, для лучшей анонимизации)
                    props.subject = ""
                    props.keywords = ""
                    props.comments = ""
                    props.last_modified_by = ""
                    props.category = ""
                    props.content_status = ""

                    # Сохраняем изменения в тот же файл
                    prs.save(file_path)
                    print(f"[OK] Очищено: {file_path}")
                    files_processed += 1

                except PythonPptxError as e:
                    print(f"[ERROR] Ошибка структуры файла {file}: {e}")
                    files_error += 1
                except Exception as e:
                    print(f"[ERROR] Не удалось обработать {file}: {e}")
                    files_error += 1

    print("\n" + "=" * 30)
    print(f"Завершено!")
    print(f"Успешно обработано: {files_processed}")
    print(f"Ошибок: {files_error}")
    print("=" * 30)


if __name__ == "__main__":
    # Введите путь к вашей папке
    target_dir = input("Введите путь к папке (для рекурсивной очистки .pptx): ").strip()

    # Убираем кавычки, если пользователь скопировал путь как "C:\Path"
    target_dir = target_dir.replace('"', '').replace("'", "")

    if os.path.isdir(target_dir):
        confirm = input(f"Внимание! Метаданные будут перезаписаны в '{target_dir}'. Продолжить? (y/n): ")
        if confirm.lower() == 'y':
            clean_pptx_metadata(target_dir)
        else:
            print("Отмена операции.")
    else:
        print("Указанный путь не является папкой или не существует.")