import sys
import argparse
from pathlib import Path
from src.cli.menu import show_main_menu
from src.core.extractor import PptxToPpc
from src.core.modifier import CompactPptxModifier
from src.core.cleaner import clean_pptx_metadata
from src.core.creator import PptxCreator  # ДОБАВЛЕНО
from src.core.errors import PptxAgentException


def main():
    parser = argparse.ArgumentParser(description="PPTX Agent Tool Suite")
    subparsers = parser.add_subparsers(dest="command", help="Доступные команды для автоматизации")

    # Команда: Извлечь (extract)
    parser_ext = subparsers.add_parser("extract", help="Извлечь текст из PPTX")
    parser_ext.add_argument("pptx_file", type=str, help="Путь к исходному .pptx")
    parser_ext.add_argument("--out", type=str, help="Путь к выходному .txt (опционально)")
    parser_ext.add_argument("--stdout", action="store_true", help="Выдать результат в консоль (для ИИ)")

    # Команда: Применить изменения (modify)
    parser_mod = subparsers.add_parser("modify", help="Применить изменения из TXT в существующий PPTX")
    parser_mod.add_argument("pptx_file", type=str, help="Путь к исходному .pptx")
    parser_mod.add_argument("txt_file", type=str, help="Путь к файлу .txt ИЛИ '-' для чтения из stdin")
    parser_mod.add_argument("--out", type=str, help="Путь для сохранения результата (опционально)")

    # --- НОВАЯ КОМАНДА: Создать (create) ---
    parser_create = subparsers.add_parser("create", help="Создать НОВУЮ презентацию из Markdown")
    parser_create.add_argument("md_file", type=str, help="Путь к .md файлу ИЛИ '-' для чтения из stdin")
    parser_create.add_argument("--out", type=str, required=True, help="Путь для сохранения нового .pptx")

    # Команда: Очистить (clean)
    parser_clean = subparsers.add_parser("clean", help="Очистить метаданные в папке")
    parser_clean.add_argument("folder", type=str, help="Путь к папке")

    args = parser.parse_args()

    if not args.command:
        try:
            show_main_menu()
        except KeyboardInterrupt:
            print("\nПрограмма прервана пользователем.")
        return

    try:
        if args.command == "extract":
            extractor = PptxToPpc(args.pptx_file)
            out_file = args.out if args.out else (None if args.stdout else args.pptx_file.replace('.pptx', '.txt'))
            text_result = extractor.extract(out_file)

            if args.stdout:
                print(text_result)
                print("SUCCESS: Данные извлечены в stdout", file=sys.stderr)
            else:
                print(f"SUCCESS: Данные извлечены в {out_file}", file=sys.stderr)

        elif args.command == "modify":
            path_obj = Path(args.pptx_file)
            out_path = args.out if args.out else str(path_obj.parent / f"modified_{path_obj.name}")
            modifier = CompactPptxModifier(args.pptx_file)

            if args.txt_file == '-':
                input_data = sys.stdin.read()
                result = modifier.apply_from_text(input_data, out_path)
            else:
                result = modifier.apply_from_file(args.txt_file, out_path)

            for warning in result["warnings"]:
                print(f"WARNING: {warning}", file=sys.stderr)
            print(f"SUCCESS: Обработано слайдов: {result['slides_processed']}. Сохранено в {out_path}", file=sys.stderr)

        # --- ОБРАБОТКА КОМАНДЫ CREATE ---
        elif args.command == "create":
            creator = PptxCreator()
            out_path = args.out

            if args.md_file == '-':
                print("Ожидание Markdown данных из stdin...", file=sys.stderr)
                input_data = sys.stdin.read()
                result = creator.create_from_text(input_data, out_path)
            else:
                result = creator.create_from_file(args.md_file, out_path)

            for warning in result.get("warnings", []):
                print(f"WARNING: {warning}", file=sys.stderr)
            print(f"SUCCESS: Создано слайдов: {result['slides_created']}. Сохранено в {out_path}", file=sys.stderr)

        elif args.command == "clean":
            processed, errors, details = clean_pptx_metadata(args.folder)
            for d in details:
                print(d, file=sys.stderr)
            print(f"SUCCESS: Очищено файлов: {processed}, Ошибок: {errors}", file=sys.stderr)

    except PptxAgentException as e:
        print(f"ERROR: Ошибка при обработке данных:\n{str(e)}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"CRITICAL ERROR: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()