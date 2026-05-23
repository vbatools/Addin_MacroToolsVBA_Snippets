**Русский** | [English](README_ENG.md)
---

# 📦 База сниппетов для надстройки Macro Tools VBA

Этот репозиторий содержит коллекцию сниппетов для надстройки [Macro Tools VBA](https://github.com/vbatools/Addin_MacroToolsVBA).

## 📖 Описание

База сниппетов предназначена для использования с надстройкой Macro Tools VBA. Сниппеты автоматически загружаются из этого репозитория в надстройку, обеспечивая быстрый доступ к часто используемым фрагментам кода VBA.

## 📁 Структура репозитория

Репозиторий организован по категориям функциональности:

| Категория | Описание | Кол-во сниппетов |
|-----------|----------|-----------------|
| [App](#app-10) | Основные процедуры и функции приложения | 10 |
| [Custom](#custom-46) | Пользовательские процедуры и функции | 46 |
| [DoLoop](#doloop-8) | Циклы (For, Do While, Do Until) | 8 |
| [FSO](#fso-9) | Операции с файловой системой | 9 |
| [Forms](#forms-16) | Элементы управления формы | 16 |
| [Http](#http-3) | HTTP запросы | 3 |
| [Iff](#iff-3) | Условные операторы | 3 |
| [LogFile](#logfile-3) | Работа с логами | 3 |
| [TxtFile](#txtfile-9) | Работа с текстовыми файлами | 9 |
| **Всего** | | **110** |

### App (10)

| Сниппет | Описание |
|---------|----------|
| `CalAu` | Calculation = xlCalculationAutomatic |
| `CalMa` | Calculation = xlCalculationManual ScreenUpdating = False |
| `Enum_` | Enum |
| `pusub` | Public Sub |
| `prsub` | Private Sub |
| `pufun` | Public Function |
| `prfun` | Private Function |
| `pucon` | Public Const |
| `ScrF` | ScreenUpdating = False Calculation = xlCalculationManual |
| `ScrT` | Calculation = xlCalculationAutomatic ScreenUpdating = True |
| `ThWb` | With ThisWorkbook.Worksheets(SHName) |
| `Wth` | With |
| `pusub` | Public Sub |
| `pufun` | Public Function |
| `pucon` | Public Const |
| `prsub` | Private Sub |
| `prfun` | Private Function |

### Custom (46)

| Сниппет | Описание |
|---------|----------|
| `addCheckBox` | Создание CheckBox в диапазоне с привязкой макроса |
| `arr` | Загрузка диапазона в массив |
| `arrDemens` | Проверка размерности массива |
| `arrFor` | Цикл по массиву |
| `arrFor2` | Цикл по 2 массивам |
| `arrSort` | Сортировка массива |
| `arrSort2DArray` | Сортировка 2D массива |
| `arrTo2D` | Преобразование 1D массива в 2D |
| `arrTranspose` | Транспонирование массива |
| `arrUniqueValuesFrom` | Уникальные значения из массива |
| `BlendColor` | Смешивание двух цветов |
| `Clear` | Очистка диапазона от значений |
| `delAllFilesInPath` | Удаление всех файлов в папке |
| `dialogFile` | Выбор файлов |
| `FileDialog` | Диалог выбора файлов |
| `extractAllFilesFromZip` | Извлечение файлов из архива |
| `fileIsBusy` | Проверка занятости файла |
| `FileHave` | Проверка наличия файла |
| `getPath` | Получение пути к папке |
| `GetListFiles` | Список файлов в папке |
| `getLetterByNomer` | Буква столбца по номеру |
| `getSheetNameCloseBook` | Имена листов из закрытой книги |
| `GetSheetNameCloseBook` | Получение имён листов из закрытой книги |
| `GetUniqueValueFromRange` | Уникальные значения из диапазона |
| `hash` | Хеш-функции |
| `HaveSheetInFile` | Проверка листа в файле |
| `LastCol` | Последний столбец |
| `LastRow` | Последняя строка |
| `Path` | Проверка пути |
| `replaceSymbols` | Замена символов в строке |
| `Resize` | Выгрузка массива в диапазон |
| `ScreenUpdatingCalculation` | Переключение ScreenUpdating/Calculation |
| `SheetCopy` | Копирование листа |
| `SheetDelete` | Удаление листа |
| `Status` | Сообщение в статус бар |
| `StatusBar` | Статус бар |
| `Status` | Статус бар |
| `wbIsOpen` | Проверка открытия книги |
| `WriterLogSub` | Логгер |

### DoLoop (8)

| Сниппет | Описание |
|---------|----------|
| `DoLoUn` | Do Loop Until |
| `DoLoWh` | Do Loop While |
| `DoUnLo` | Do Until Loop |
| `DoWhLo` | Do While Loop |
| `ForEa` | For Each ... In |
| `ForEaW` | For Each лист в книге |
| `ForI` | For I = 1 To 10 |
| `ForIS` | For I = 1 To 10 Step 1 |

### FSO (9)

| Сниппет | Описание |
|---------|----------|
| `delPath` | Удаление папки |
| `MoveFile` | Перемещение файла |
| `sFileExists` | Проверка наличия файла |
| `sFolderExists` | Проверка наличия папки |
| `sGetBaseName` | Имя файла без расширения |
| `sGetExtensionName` | Расширение файла |
| `sGetFileName` | Имя файла |
| `sGetParentFolderName` | Родительская папка |
| `sGetTempName` | Случайное имя файла |

### Forms (16)

| Сниппет | Описание |
|---------|----------|
| `clsAnchors` | Резиновая форма |
| `clsButtonIcon` | Кнопка с иконкой |
| `clsCalendarDate` | Календарь для TextBox |
| `clsCalendarTime` | Время для TextBox |
| `clsComboBox` | ComboBox Multi класс |
| `clsComboBoxForm` | ComboBox с модальным окном |
| `clsContextMenu` | Контекстное меню |
| `cls ListBox` | ListBox класс |
| `clsModernStyle` | Стиль UserForms |
| `clsProgresBar` | Прогресс бар |
| `clsSlider` | Слайдер |
| `clsTextArea` | TextArea контрол |
| `clsTextboxMask` | TextBox с маской |
| `fint` | Центрирование формы |
| `ftxtbtn` | DropButton для TextBox |
| `ftxtbtnclick` | Клик DropButton |

### Http (3)

| Сниппет | Описание |
|---------|----------|
| `responseTextHttp` | HTTP запрос |
| `webPageHTML` | HTML страница |
| `webPageText` | Текст HTML страницы |

### Iff (3)

| Сниппет | Описание |
|---------|----------|
| `IfEl` | IF Else |
| `IfEn` | IF |
| `SeCa` | Select Case |

### LogFile (3)

| Сниппет | Описание |
|---------|----------|
| `ShowLog` | Показать лог |
| `WriteErrorLog` | Лог ошибок |
| `WriteLog` | Лог запись |

### TxtFile (9)

| Сниппет | Описание |
|---------|----------|
| `Base64` | Base64 кодирование/декодирование |
| `base64ToFile` | Base64 в файл |
| `fileToBase64` | Файл в Base64 |
| `loadTextFromTextFile` | Загрузка текста из файла |
| `Range2CSV` | Диапазон в CSV |
| `saveTextToFile` | Сохранение текста в файл |
| `TXTAddIntoTXTFile` | Добавление в файл |
| `TXTReadALLFile` | Чтение всего файла |
| `TXTWriteFile` | Запись в файл |

## 🚀 Использование

Для использования сниппетов установите надстройку [Macro Tools VBA](https://github.com/vbatools/Addin_MacroToolsVBA) и настройте её на использование этого репозитория как источника сниппетов.

## 👤 Автор

**VBATools**

## 📄 Лицензия

[Apache License 2.0](LICENSE)