![](images/ru.png) **Russian** | ![](images/en.png) [English](readme.md)

# Excel-надстройка Navferty's Tools

![Navferty's Tools Ribbon Tab in MS Excel](images/NavfertyToolsRibbonRu.png)

## Навигация
  - [Управление](#управление)
    - [Отменить последнее действие](#отменить-последнее-действие)
  - [Преобразование значений](#преобразование-значений)
    - [Парсинг числовых значений](#парсинг-числовых-значений)
    - [Заменить](#заменить)
      - [Транслитерацией](#транслитерацией)
      - [На аналоги](#на-аналоги)
    - [Преобразовать в слова](#преобразовать-в-слова)
  - [Форматирование значений](#форматирование-значений)
    - [Изменить регистр](#изменить-регистр)
    - [Удалить пробелы](#удалить-пробелы)
  - [Общие](#общие)
    - [Снять защиту с книги и листов](#снять-защиту-с-книги-и-листов)
    - [Множественная подсветка дублей](#множественная-подсветка-дублей)
    - [Разъединение ячеек с заполнением](#разъединение-ячеек-с-заполнением)
    - [Найти все ошибки в выделенном диапазоне](#найти-все-ошибки-в-выделенном-диапазоне)
    - [Копировать в Markdown](#копировать-в-markdown)
    - [Валидация значений](#валидация-значений)
  - [Работа с XML](#работа-с-xml)
    - [Создать образец XML (по XSD)](#создать-образец-xml-по-xsd)
    - [Валидация XML по XSD](#валидация-xml-по-xsd)

  - [В разработке](#в-разработке)
    - [Сократить название организаций](#сократить-название-организаций)

  - [Инструкция по установке](#инструкция-по-установке)

---

## Управление

### Отменить последнее действие

|||
|:-:|---|
|![](images/icons/undo.png)|Отменить последнее действие, совершённое с помощью данной надстройки. Отмена возможна для некоторых функций в разделах 'Преобразование значений' и 'Форматирование значений' и только, если диапазон ячеек не редактировался после совершения действия.|

[Наверх](#excel-надстройка-navfertys-tools)

---

## Преобразование значений

### Парсинг числовых значений

|||
|:-:|---|
|![](images/icons/parseNumerics.png)|Преобразование чисел, сохраненных в виде текста, в числовой формат.|

<details>
  <summary>Посмотреть скриншоты</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/ParseNumerics1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ParseNumerics2.png)
</details>

### Заменить

|||
|:-:|---|
|![](images/icons/replace.png)|Замена русских символов по таблицам соответствия.|

*   #### Транслитерацией

    |||
    |:-:|---|
    |![](images/icons/transliterate.png)|Полностью меняется весь русский алфавит на английский. Например, буква "Ж" заменится на "Zh", а буква "Щ" - на "Shch". Основано на документе ИКАО [9303](https://www.icao.int/publications/Documents/9303_p3_cons_ru.pdf).|

    <details>
      <summary>Посмотреть скриншоты</summary>

    ![Navferty's Tools Ribbon Tab in MS Excel](images/Transliterate1.png)
    ![Navferty's Tools Ribbon Tab in MS Excel](images/Transliterate2.png)
    </details>

* #### На аналоги

    |||
    |:-:|---|
    |![](images/icons/replaceChars.png)|Заменятся только схожие по верхнему регистру буквы алфавитов, такие как: Аа, Вв, Ее, Кк, Мм, Нн, Оо, Рр, Сс, Тт, Уу, Хх.|

    <details>
      <summary>Посмотреть скриншоты</summary>

    ![Navferty's Tools Ribbon Tab in MS Excel](images/ReplaceChars1.png)
    ![Navferty's Tools Ribbon Tab in MS Excel](images/ReplaceChars2.png)
    </details>

### Преобразовать в слова

|||
|:-:|---|
|![](images/icons/stringifyNumerics.png)|Переписывает числовые значения в текст с расшифровкой<br>- На русском<br>- На английском<br>- На французском|

<details>
  <summary>Посмотреть скриншоты</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Stringify1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/Stringify2.png)
</details>

[Наверх](#excel-надстройка-navfertys-tools)

---

## Форматирование значений

### Изменить регистр

|||
|:-:|---|
|![](images/icons/toggleCase.png)|Переключение регистра для текстовых значений в выделенных ячейках по схеме:<br>`Абвгд` -> `абвгд` -> `АБВГД`|

<details>
  <summary>Посмотреть скриншоты</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase2.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase3.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase4.png)
</details>

### Удалить пробелы

|||
|:-:|---|
|![](images/icons/trimSpaces.png)|Очистить текстовое содержимое выделенных ячеек от лишних пробелов. Удаляет повторные пробелы и переносы строк, а также начальные и конечных пробелы в ячейках, имеющих текстовый формат.|

<details>
  <summary>Посмотреть скриншоты</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/TrimSpaces1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/TrimSpaces2.png)
</details>

[Наверх](#excel-надстройка-navfertys-tools)

---

## Общие

### Снять защиту с книги и листов

|||
|:-:|---|
|![](images/icons/unprotectWorkbook.png)|Позволяет снять защиту со всех листов открытой книги, как и защиту всей книги, в отсутствие пароля, а также разблокировать VBA проект (при его наличии), на который установлен пароль. Функция не распространяется на зашифрованные книги.|

### Множественная подсветка дублей

|||
|:-:|---|
|![](images/icons/highlightDuplicates.png)|Устанавливает цвет ячеек, которые содержат повторяющиеся значения, в выделенном диапазоне. Различные цвета соответствуют разным группам дублей.|

<details>
  <summary>Посмотреть скриншот</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Duplicates.png)
</details>

### Разъединение ячеек с заполнением

|||
|:-:|---|
|![](images/icons/unmergeCells.png)|Отменяет объединение ячеек, с одновременным заполнением разъединенных ячеек исходым значением.|

<details>
  <summary>Посмотреть скриншоты</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Unmerge1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/Unmerge2.png)
</details>

### Найти все ошибки в выделенном диапазоне

|||
|:-:|---|
|![](images/icons/findErrors.png)|Поиск всех ячеек в выделенных ячейках, которые содержат ошибки вычислений:<br><br>Виды ошибок в MS Excel:<br>`#Н/Д`<br>`#ИМЯ?`<br>`#ДЕЛ/0!`<br>`#ССЫЛКА!`<br>`#ЗНАЧ!`<br>`#ЧИСЛО!`<br>`#ПУСТО!`|

<details>
  <summary>Посмотреть скриншот</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/FindErrorValues.png)
</details>

### Копировать в Markdown

|||
|:-:|---|
|![](images/icons/parseNumerics.png)|Копировать в буфер обмена выделенный фрагмент как таблица Markdown.|

<details>
  <summary>Посмотреть скриншот</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/ExportToMarkdown.png)
</details>

### Валидация значений

|||
|:-:|---|
|![](images/icons/validation.png)|<p>Проверить значения ячеек в выделенном диапазоне на соответствие определённому формату.<br><br>Поддерживаемые форматы:<br>- Число<br>- Дата<br>- ИНН ФЛ\* (12 цифр, с двумя проверочными цифрами)<br>- ИНН ЮЛ\* (10 цифр, с одной проверочной цифрой)<br>- Текст для XML (отсутствие символов `<` и `>` и иных невалидных для содержимого XML)<br><br>\* _- Корректность ИНН не гарантирует существование организации или физического лица, которому бы принадлежал этот ИНН_</p>|

<details>
  <summary>Посмотреть скриншот</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Validate.png)
</details>

[Наверх](#excel-надстройка-navfertys-tools)

---

## Работа с XML

### Создать образец XML (по XSD)

|||
|:-:|---|
|![](images/icons/createSampleXml.png)|Сгенерировать пример XML-файла по xsd-схеме. Позволяет получить наглядное представление о структуре xml-файла при наличии XSD-файла со схемой.|

Например, для приведённой ниже схемы
<details>
  <summary>Образец схемы - sample.xsd</summary>

```xml
<?xml version="1.0" encoding="UTF-8" ?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">

<xs:element name="КорневойЭлемент">
  <xs:complexType>
    <xs:sequence>
      <xs:element name="Заголовок" type="xs:string"/>
      <xs:element name="Адресат">
        <xs:complexType>
          <xs:sequence>
            <xs:element name="Имя" type="xs:string"/>
            <xs:element name="Адрес" type="xs:string"/>
            <xs:element name="Город" type="xs:string"/>
            <xs:element name="Страна" type="xs:string"/>
          </xs:sequence>
        </xs:complexType>
      </xs:element>
      <xs:element name="Элемент" maxOccurs="3">
        <xs:complexType>
          <xs:sequence>
            <xs:element name="Название" type="xs:string"/>
            <xs:element name="Описание" type="xs:string" minOccurs="0"/>
            <xs:element name="Количество" type="xs:positiveInteger"/>
            <xs:element name="Цена" type="xs:decimal"/>
          </xs:sequence>
        </xs:complexType>
      </xs:element>
    </xs:sequence>
    <xs:attribute name="Идентификатор" type="xs:string" use="required"/>
  </xs:complexType>
</xs:element>

</xs:schema>
```
</details>

Будет сформирован такой xml-файл:

<details>
  <summary>Результат генерации - sample.xml</summary>

```xml
<КорневойЭлемент xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" Идентификатор="Идентификатор1">
  <Заголовок>Заголовок1</Заголовок>
  <Адресат>
    <Имя>Имя1</Имя>
    <Адрес>Адрес1</Адрес>
    <Город>Город1</Город>
    <Страна>Страна1</Страна>
  </Адресат>
  <Элемент>
    <Название>Название1</Название>
    <Описание>Описание1</Описание>
    <Количество>1</Количество>
    <Цена>1</Цена>
  </Элемент>
  <Элемент>
    <Название>Название2</Название>
    <Описание>Описание2</Описание>
    <Количество>79228162514264337593543950335</Количество>
    <Цена>-79228162514264337593543950335</Цена>
  </Элемент>
  <Элемент>
    <Название>Название3</Название>
    <Описание>Описание3</Описание>
    <Количество>2</Количество>
    <Цена>79228162514264337593543950335</Цена>
  </Элемент>
</КорневойЭлемент>
```
</details>

### Валидация XML по XSD

|||
|:-:|---|
|![](images/icons/validateXml.png)|Проверить соответствие xml-файла на основе XSD-схемы. Список ошибок будет выведен на лист в новой книге MS Excel.|

Пример отчёта о проверке:

|Уровень|Элемент|Ошибка|
|---|---|---|
|Error|Город|Элемент "Адресат" имеет недопустимый дочерний элемент "Город". Список ожидаемых элементов: "Адрес".|
|Error|Количество|Элемент "Количество" недействителен: значение "-5" недействительно с точки зрения его типа данных "http://www.w3.org/2001/XMLSchema:positiveInteger" — Значение "-5" недопустимо мало или недопустимо велико для PositiveInteger.|
|Error|Цена|Элемент "Цена" недействителен: значение "не число" недействительно с точки зрения его типа данных "http://www.w3.org/2001/XMLSchema:decimal" — Строка "не число" не является допустимым значением Decimal.|

[Наверх](#excel-надстройка-navfertys-tools)

---

## В разработке

### Сократить название организаций

Позволяет сократить тип организации до аббревиатуры (ООО, ПАО, ОАО и др.).

[Наверх](#excel-надстройка-navfertys-tools)

---

## Инструкция по установке

Решение построено в Azure, вы можете скачать установочные файлы оттуда:

* Посетите https://navferty.visualstudio.com/NavfertyExcelAddIn/_build?definitionId=3

* Выберите последнюю версию сборки 'NavfertyExcelAddIn - Publish':

  ![Navferty's Tools Ribbon Tab in MS Excel](images/Install1.png)

* Загрузите опубликованные установочные файлы:

  ![Navferty's Tools Ribbon Tab in MS Excel](images/Install2.png)

* Извлеките файлы в папку и запустите файл '.vsto':

  > Настоятельно рекомендуется использовать папку рабочего стола - установка обновлений разрешена только из той папки, из которой надстройка была установлена в первый раз!

  ![Navferty's Tools Ribbon Tab in MS Excel](images/Install3.png)

* После завершения процесса установки запустите (или перезагрузите) приложение Excel, и вы увидите новую вкладку:

  ![Navferty's Tools Ribbon Tab in MS Excel](images/Install4.png)

[Наверх](#excel-надстройка-navfertys-tools)