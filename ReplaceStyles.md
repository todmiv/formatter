Метод `ReplaceStyles` из Open XML SDK — это официальный пример от Microsoft, который полностью заменяет все стили в целевом документе на стили из документа-источника. Он работает на уровне частей (parts) DOCX-файла, что делает его надежным и контролируемым решением.

Ниже — подробное руководство по его использованию.

### ⚙️ Требования и подготовка

Перед использованием кода необходимо выполнить несколько шагов.

1.  **Установить Open XML SDK**: В вашем проекте должен быть установлен пакет NuGet [`DocumentFormat.OpenXml`](https://www.nuget.org/packages/DocumentFormat.OpenXml/).
2.  **Добавить ссылки на сборки**: В проект необходимо добавить ссылки на следующие сборки:
    *   `WindowsBase`
    *   `DocumentFormat.OpenXml`
3.  **Добавить директивы `using`**: В начало файла с кодом добавьте следующие пространства имен:
    ```csharp
    using System.IO;
    using System.Xml;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Packaging;
    ```

### 🧠 Как это работает

Метод `ReplaceStyles` работает в два основных этапа.

1.  **Извлечение (`ExtractStylesPart`)**: Этот вспомогательный метод открывает документ-источник, находит в нем часть со стилями (`styles.xml` или `stylesWithEffects.xml`) и извлекает ее содержимое в объект `XDocument`.
2.  **Замена (`ReplaceStylesPart`)**: Этот метод принимает извлеченный `XDocument` и целевому документу. Он находит в целевом документе соответствующие части стилей и полностью заменяет их содержимое на новое из `XDocument`.

Важно отметить, что код обрабатывает **оба** возможных варианта частей со стилями: стандартную (`styles`) и дополнительную (`stylesWithEffects`), которая появляется в более новых версиях Word.

### 💻 Полный пример кода

Ниже приведен полный рабочий код, основанный на официальном примере от Microsoft.

```csharp
using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

class Program
{
    // Главный метод, который вы вызываете из своей программы
    static void ReplaceStyles(string fromDoc, string toDoc)
    {
        // 1. Извлекаем и заменяем стандартную часть стилей
        XDocument? stylesPart = ExtractStylesPart(fromDoc, false);
        if (stylesPart is not null)
        {
            ReplaceStylesPart(toDoc, stylesPart, false);
        }

        // 2. Извлекаем и заменяем дополнительную часть стилей (для новых версий Word)
        XDocument? stylesWithEffectsPart = ExtractStylesPart(fromDoc, true);
        if (stylesWithEffectsPart is not null)
        {
            ReplaceStylesPart(toDoc, stylesWithEffectsPart, true);
        }
    }

    // Вспомогательный метод для извлечения части стилей
    static XDocument? ExtractStylesPart(string fileName, bool isStylesWithEffectsPart)
    {
        // Открываем документ только для чтения
        using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
        {
            // Находим нужную часть
            StylesPart? stylesPart = null;
            if (isStylesWithEffectsPart)
            {
                // Ищем часть StylesWithEffects
                stylesPart = document.MainDocumentPart?.StylesWithEffectsPart;
            }
            else
            {
                // Ищем стандартную часть Styles
                stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
            }

            // Если часть найдена, загружаем её XML в XDocument и возвращаем
            if (stylesPart is not null)
            {
                using (StreamReader reader = new StreamReader(stylesPart.GetStream()))
                {
                    return XDocument.Load(reader);
                }
            }
            return null;
        }
    }

    // Вспомогательный метод для замены части стилей
    static void ReplaceStylesPart(string fileName, XDocument newStyles, bool isStylesWithEffectsPart)
    {
        // Открываем целевой документ для редактирования
        using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
        {
            // Находим часть, которую нужно заменить
            StylesPart? stylesPart = null;
            if (isStylesWithEffectsPart)
            {
                stylesPart = document.MainDocumentPart?.StylesWithEffectsPart;
            }
            else
            {
                stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
            }

            // Если часть найдена, заменяем её содержимое
            if (stylesPart is not null)
            {
                using (StreamWriter writer = new StreamWriter(stylesPart.GetStream(FileMode.Create)))
                {
                    newStyles.Save(writer);
                }
            }
        }
    }

    // Пример использования
    static void Main(string[] args)
    {
        // Путь к файлу, из которого берутся стили
        string sourceDocument = @"C:\Путь\К\Документу-Источнику.docx";
        // Путь к файлу, в котором нужно заменить стили
        string targetDocument = @"C:\Путь\К\Целевому-Документу.docx";

        // Вызов метода для замены стилей
        ReplaceStyles(sourceDocument, targetDocument);

        Console.WriteLine("Стили успешно заменены.");
    }
}
```

### 📝 Важные замечания

*   **Необратимость**: Этот метод полностью **заменяет**, а не объединяет стили. Все существующие стили в целевом документе будут удалены и заменены стилями из документа-источника.
*   **Внешний вид**: После выполнения кода весь текст в целевом документе, который использует стили, изменит свой внешний вид в соответствии с новыми стилями из документа-источника.
*   **Обработка ошибок**: Для промышленного использования код следует дополнить обработкой исключений (например, на случай, если файлы не найдены или повреждены).

Этот подход является наиболее прямым и надежным для полной замены стилей, так как он работает на уровне структуры самого DOCX-файла.

---

## Реализация на Python (в проекте Formatter_modification)

Аналогичный подход реализован на Python в `src/apply/style_applier.py`:

```python
from src.apply.style_applier import StyleApplier

applier = StyleApplier("template.docx")
applier.apply("target.docx", output_path="output.docx", replace_styles_raw=True)
```

**CLI:**
```bash
python -m src.apply.style_applier --template template.docx --target target.docx --replace-raw
```

**GUI:** вкладка "Замена стилей" в главном окне приложения.

Метод `_replace_styles_raw()` работает через `zipfile` — извлекает `word/styles.xml` и `word/stylesWithEffects.xml` из DOCX-источника и заменяет их в целевом документе.