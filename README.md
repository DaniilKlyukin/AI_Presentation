# PPTX Tool Suite

Мощный инструмент на Python для автоматизации создания и редактирования презентаций PowerPoint. Основным компонентом является **Creator** — модуль для генерации стильных презентаций из обычных Markdown-файлов.

## 🌟 Основные возможности (Creator)

*   **Markdown -> PPTX**: Превращайте структурированный текст в презентацию за секунды.
*   **Автоматическое разбиение**: Если текст не помещается на один слайд, система автоматически создаст новый слайд с пометкой "(продолжение)".
*   **LaTeX Формулы**: Поддержка математических выражений. Простые формулы конвертируются в Unicode, сложные — рендерятся в высококачественные PNG через Matplotlib.
*   **Подсветка кода**: Вставка блоков кода с сохранением синтаксиса (используется Pygments).
*   **Гибкая настройка**: Полный контроль над цветами, шрифтами, отступами и нумерацией через файл `.env`.
*   **Таблицы и изображения**: Поддержка стандартной разметки Markdown для вставки визуального контента.

---

## 📸 Демонстрация

| Входной Markdown | Результат в PPTX |
| :--- | :--- |
| ![Markdown Example](https://github.com/user-attachments/assets/1b11e5f8-6b98-4cca-9c88-47b5f0b6c0b2) | ![PPTX Result](https://github.com/user-attachments/assets/712861b6-2f64-420d-aba8-92c5d1b38a69) |
| ![Markdown Example](https://github.com/user-attachments/assets/2599cc77-d390-42e6-9190-596dd29b0ace) | ![PPTX Result](https://github.com/user-attachments/assets/1b262913-697d-49d6-a02e-87d1d99c154c) |
| ![Markdown Example](https://github.com/user-attachments/assets/ee68f5d9-25b9-4b7a-b484-c4fb752dae35) | ![PPTX Result](https://github.com/user-attachments/assets/f89ab59c-e9c2-45ae-84b5-6079250e8411) |
| ![Markdown Example](https://github.com/user-attachments/assets/f6d08854-d55b-4fe4-b71e-2d59774108a5) | ![PPTX Result](https://github.com/user-attachments/assets/79b381bf-b01b-410a-848e-051b4feb1300) |
| ![Markdown Example](https://github.com/user-attachments/assets/54faed4c-941b-4b0b-a242-755501fc96db) | ![PPTX Result](https://github.com/user-attachments/assets/b0ba4cf2-239b-42b9-913f-776aa72c8b10) |

---

## 🚀 Быстрый старт

### Установка

1. Клонируйте репозиторий.
2. Установите зависимости:
```bash
pip install python-pptx marko matplotlib pygments python-dotenv pylatexenc
```

### Создание презентации
Подготовьте файл `presentation.md`. Используйте `---` для разделения слайдов. Первый блок всегда становится титульным слайдом.

```bash
python main.py create presentation.md --out my_presentation.pptx
```

---

## 🛠 Другие инструменты (Legacy)

В состав также входят инструменты для работы с ИИ-агентами (экспериментально):
*   **Extract**: Выгружает структуру и текст PPTX в компактный текстовый формат для чтения нейросетью.
*   **Modify**: Позволяет применить текстовые правки от ИИ обратно в PPTX (работает в режиме патча).
*   **Clean**: Массовая очистка метаданных (автор, тема, комментарии) во всех презентациях в папке.

*Примечание: Инструменты модификации могут работать нестабильно на сложных макетах.*

---

## ⚙️ Конфигурация (.env)

Настройте внешний вид презентации под свой корпоративный стиль без правки кода. Создайте файл `.env` в корне проекта:

### Основные настройки
* `PPT_ASPECT_RATIO`: Соотношение сторон (16:9 или 4:3).
* `PPT_FONT`: Основной шрифт (например, "Bookman Old Style").
* `PPT_BODY_SIZE`: Размер основного текста.
* `PPT_LINE_SPACING`: Межстрочный интервал.

### Заголовок (Title Bar)
* `PPT_TITLE_BG_COLOR`: Цвет фона плашки заголовка в формате `R,G,B`.
* `PPT_TITLE_FONT_COLOR`: Цвет текста заголовка.
* `PPT_TITLE_HEIGHT_CM`: Высота области заголовка.

### Колонтитулы и нумерация
* `PPT_FOOTER_TEXT`: Текст в нижней части слайда.
* `PPT_SLIDE_NUMBERING`: Включить нумерацию в формате `X/Total` (True/False).
* `PPT_FOOTER_BORDER_COLOR`: Цвет разделительной линии.

### Математика и Код
* `PPT_FORMULA_NUMBERING`: Автоматическая нумерация выносных формул.
* `PPT_CODE_FONT`: Шрифт для блоков кода.
* `PPT_CODE_SIZE`: Размер шрифта в блоках кода.

---

## 📝 Пример Markdown для генерации

```markdown
# Моя крутая презентация
Автор: Иван Иванов
2024

---

# Заголовок контентного слайда
* Первый пункт списка
* Второй пункт с **жирным текстом**
* Формула inline: $E = mc^2$

$$
\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}
$$

---

# Слайд с кодом
```python
def hello_world():
    print("Hello from PPTX Creator!")
```
```

---

## 📄 Лицензия
MIT. Используйте и модифицируйте как угодно!
