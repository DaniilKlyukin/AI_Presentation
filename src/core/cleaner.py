import os
from pptx import Presentation
from pptx.exc import PythonPptxError

def clean_pptx_metadata(folder_path):
    """
    Рекурсивно обходит папку и очищает метаданные во всех pptx файлах.
    Возвращает кортеж (processed_count, error_count, details_list).
    """
    files_processed = 0
    files_error = 0
    details = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".pptx") and not file.startswith("~$"):
                file_path = os.path.join(root, file)

                try:
                    prs = Presentation(file_path)
                    props = prs.core_properties

                    props.title = ""
                    props.author = ""
                    props.subject = ""
                    props.keywords = ""
                    props.comments = ""
                    props.last_modified_by = ""
                    props.category = ""
                    props.content_status = ""

                    prs.save(file_path)
                    files_processed += 1
                    details.append(f"[OK] Очищено: {file_path}")

                except PythonPptxError as e:
                    files_error += 1
                    details.append(f"[ERROR] Ошибка структуры файла {file}: {e}")
                except Exception as e:
                    files_error += 1
                    details.append(f"[ERROR] Не удалось обработать {file}: {e}")

    return files_processed, files_error, details