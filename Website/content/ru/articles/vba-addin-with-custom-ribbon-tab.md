{
  "title": "Создание VBA-надстройки с отдельной вкладкой в MS Excel",
  "description": "Как упаковать VBA-макросы в надстройку .xlam с собственной вкладкой на ленте: правка книги как zip-архива, customUI.xml, обработчики и установка.",
  "image": "https://navferty.net/images/articles/vba-addin-with-custom-ribbon-tab/1578783854145094499.png"
}

[← Все статьи](/ru/articles/)

# Создание VBA-надстройки с отдельной вкладкой в MS Excel

Как упаковать VBA-макросы в надстройку .xlam с собственной вкладкой на ленте: правка книги как zip-архива, customUI.xml, обработчики и установка.

Многие, кто часто работает в Excel, накопили коллекцию VBA-макросов, облегчающих ежедневный труд. Одни хранят макросы в отдельной книге, другие собрали их в <a href="https://support.office.com/en-us/article/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566" rel="noopener" target="_blank">«личной книге макросов»</a> (personal.xlsb), доступной на уровне всего приложения, и вручную добавили кнопки вызова на панель инструментов. В первом случае коллекцией удобно делиться с коллегами (достаточно переслать файл), но чтобы ею воспользоваться, книгу нужно каждый раз открывать. Во втором случае доступ к функциям есть сразу при запуске Excel, но с передачей наработок другим пользователям могут быть сложности.

В этой статье описан способ сделать VBA-надстройку с собственной вкладкой на ленте Excel (*Ribbon*), которая сочетает преимущества обоих подходов. Это файл в формате xlam, которым можно делиться как обычной книгой. Когда пользователь открывает его, на ленте появляется вкладка с доступом к функциям надстройки:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578783854145094499.png" loading="lazy" width="700" height="108" alt="Новая вкладка Sample Add-in на ленте Excel" />

При желании надстройку можно установить на постоянной основе, чтобы не открывать её каждый раз.

Понадобятся: Excel, 7-Zip и Visual Studio Code. Последние два необязательны: подойдёт любой архиватор с поддержкой zip и любой текстовый редактор, умеющий сохранять в UTF-8 (если вы хотите использовать кириллицу). Желательно также иметь базовое представление о формате XML.

Код и готовую книгу XLSM можно найти в [репозитории SampleExcelXlamAddin на GitHub](https://github.com/navferty/SampleExcelXlamAddin).

## Подготовка книги

Откроем новую книгу и добавим модуль с образцом кода:

    Option Explicit

    Public Sub SayHelloWorld()
        Dim firstCellValue As String

        firstCellValue = ThisWorkbook.Worksheets(1).Cells(1, 1).Value
        MsgBox "Hello, world! Cell A1 value is " & firstCellValue
    End Sub

Сохраним книгу в формате xlsm (книга с поддержкой макросов). Современные форматы документов Office (*xlsx, xlsm, docx, docm, pptx* и т. д.) основаны на стандарте OpenXML и представляют собой обычный zip-архив из компонентов (XML-файлы листов, изображения, двоичное содержимое и прочее) и файлов отношений (*.rels*), которые задают структуру документа ссылками на компоненты.

Чтобы увидеть содержимое книги, можно переименовать *Sample.xlsm* в *Sample.zip* и открыть любым архиватором. Некоторые архиваторы, например 7-Zip, умеют открывать такие файлы напрямую, без переименования:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578771762187878642.webp" loading="lazy" width="700" height="492" alt="Содержимое xlsm-файла в архиваторе" />

Чтобы добавить панель на ленту, нужно создать новый компонент **customUI/customUI.xml** и добавить ссылку на него в корневой файл отношений **\_rels/.rels**.

Распакуем весь документ в отдельную папку (в 7-Zip: «Распаковать в "Sample\\») и создадим в ней папку **customUI**:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/15787721721968742.webp" loading="lazy" width="700" height="394" alt="Папка customUI в распакованной книге" />

Теперь добавим сам компонент:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578772671166267580.webp" loading="lazy" width="700" height="427" alt="Файл customUI.xml в папке customUI" />

## Файл customUI.xml

Текстовая версия содержимого **customUI.xml**:

    <?xml version="1.0" encoding="utf-8"?>
    <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
      <ribbon>
        <tabs>
          <tab id="customTab" label="Sample Add-in" keytip="XT">
            <group id="Tools" label="Group 1">
              <button id="HelloWorldBtn" imageMso="BlackAndWhiteAutomatic" keytip="H" label="Привет, Мир!"
                      onAction="SayHelloWorld" screentip="Поприветствовать мир" size="large" />
              <button id="DuplicatesBtn" imageMso="SmartArtChangeColorsGallery" keytip="D"
                      label="Выделение цветом дублей" onAction="DuplicateColors"
                      screentip="Выделение парными цветами повторов в выделенном диапазоне" size="large" />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>

Коротко о важных элементах и атрибутах.

- Значения идентификаторов (атрибут *id*) должны быть уникальными. Для некоторых элементов (например, *tab*, когда вы объявляете новую вкладку, а не существующую через *idMso*) нужно обязательно указать и *label*, иначе Excel не покажет элемент.
- Атрибут *imageMso* задаёт значок элемента из числа имеющихся в Office. Список доступных значений большой, но легко находится поиском.
- Элемент *button* — обычная кнопка. В зависимости от атрибута *size* она может быть маленькой или большой, во всю высоту вкладки.
- Атрибут *onAction* задаёт имя процедуры, которая будет выполнена при основном действии с элементом (например, при нажатии кнопки). У процедуры должна быть определённая сигнатура (число и тип аргументов), её можно найти в спецификации. Например, обработчик *onAction* для button принимает один аргумент типа `IRibbonControl`.
- Имя процедуры не должно совпадать с именем модуля (или нужно явно указать модуль перед именем: `Module1.SayHelloWorld`).
- Атрибут *keytip* помогает при навигации с клавиатуры: клавишей Alt можно переходить по вкладкам Excel.

Образец разметки есть в документации Microsoft: <a href="https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-add-custom-ui-to-a-spreadsheet-document" rel="noopener" target="_blank">образец customUI</a>. Тем, кто хочет изучить стандарт подробно, пригодится спецификация: <a href="https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43" rel="noopener" target="_blank">[MS-CUSTOMUI]: Custom UI XML Markup Specification</a>.

Если в файле есть не-ASCII символы (например, кириллица), сохраняйте его в кодировке UTF-8, иначе Excel может не показать новую вкладку.

## Ссылка в файле .rels

Чтобы Excel «увидел» новый компонент и понял, что это, нужно сослаться на него в корневом файле отношений в папке **\_rels**, он называется **.rels**. Это тоже XML, и для удобства его можно отформатировать (в VS Code: Ctrl+Shift+P → Format Document).

Добавляем элемент *Relationship* с путём к customUI.xml, типом компонента и идентификатором (какой именно, не важно, главное, чтобы он был уникальным среди других элементов *Relationship*):

    <Relationship Id="rID4" Target="customUI/customUI.xml" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" />

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578772992128329742.webp" loading="lazy" width="700" height="427" alt="Файл .rels с добавленным элементом Relationship" />

Теперь запакуйте все файлы обратно в архив и верните имя Sample.xlsm.

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878107816842521.webp" loading="lazy" width="700" height="394" alt="Упаковка файлов обратно в архив" />

Если всё сделано правильно, при открытии книги на ленте появится новая вкладка:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578781539166930736.png" loading="lazy" width="700" height="256" alt="Новая вкладка на ленте после открытия книги" />

## Ошибка «Wrong number of arguments»

При запуске макроса HelloWorld можно увидеть предупреждение:

    Wrong number of arguments or invalid property assignment

Причина в том, что процедура-обработчик должна иметь определённую сигнатуру. Обработчик нажатия на button выглядит так:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578781698127844347.webp" loading="lazy" width="700" height="332" alt="Сигнатура обработчика нажатия кнопки" />

Добавьте аргумент `rc As IRibbonControl`, и появится долгожданное приветствие:

    Option Explicit

    Public Sub SayHelloWorld(rc As IRibbonControl)
        Dim firstCellValue As String

        firstCellValue = ThisWorkbook.Worksheets(1).Cells(1, 1).Value
        MsgBox "Hello, world! Cell A1 value is " & firstCellValue
    End Sub

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878173618865220.png" loading="lazy" width="211" height="133" alt="Окно с приветствием после нажатия кнопки" />

Список сигнатур есть на <a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)#how-can-i-determine-the-correct-signatures-for-each-callback-procedure" rel="noopener" target="_blank">сайте документации Microsoft</a>.

## Сохранение как надстройки

Осталось сохранить книгу с макросами как надстройку. В редакторе VBE выберите объект `ThisWorkbook` и установите свойство `IsAddin` в True.

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578782774119989254.webp" loading="lazy" width="700" height="445" alt="Свойство IsAddin объекта ThisWorkbook" />

Затем сохраните книгу в новом формате (Excel предложит это сам, если нажать Ctrl+S):

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878283811024502.webp" loading="lazy" width="700" height="394" alt="Сохранение книги в формате надстройки" />

Теперь надстройку можно открыть как обычную книгу. Отдельным окном она не отображается, но видна в Project Explorer редактора VBE:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878307317991772.webp" loading="lazy" width="700" height="497" alt="Надстройка в Project Explorer редактора VBA" />

Чтобы надстройка была доступна сразу при запуске Excel, установите её через Excel Add-ins на вкладке «Разработчик»...

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578783265155493627.webp" loading="lazy" width="700" height="479" alt="Установка надстройки через вкладку «Разработчик»" />

...или через параметры Excel, раздел «Надстройки» (Add-ins):

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878334215840375.webp" loading="lazy" width="700" height="574" alt="Установка надстройки через настройки Excel" />

Готово. Теперь надстройкой можно пользоваться при работе с любым документом, а делиться наработками с коллегами легко, просто пересылая файл надстройки.

## Бонус: другие элементы Custom UI

- **toggleButton** — кнопка с двумя состояниями;
- **splitButton** — составной элемент из button или toggleButton и выпадающего меню;
- **dropDown** — выпадающий список с заранее заданным набором элементов;
- **comboBox** — поле ввода, которое может содержать заранее заданный набор элементов;
- **dynamicMenu** — выпадающий список, элементы которого определяются динамически в методе getContent;
- **checkBox** — флажок с двумя состояниями.

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578831453120134147.png" loading="lazy" width="700" height="108" alt="Примеры элементов Custom UI на ленте" />

Разметка и код модуля VBA приведены ниже, их же можно найти в [репозитории SampleExcelXlamAddin на GitHub](https://github.com/navferty/SampleExcelXlamAddin).

**customUI.xml**

    <?xml version="1.0" encoding="utf-8"?>
    <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
      <ribbon>
        <tabs>
          <tab id="customTab" label="Sample Add-in" keytip="XT">
            <group id="SampleControls" label="Group 2">
              <splitButton id="splitButton" size="large">
                <button id="SplitButton" imageMso="HappyFace" label="Split Button" />
                <menu id="menu">
                  <button id="button1" label="Button 1" onAction="OnSplitButton1Click" />
                  <button id="button2" label="Button 2" onAction="OnSplitButton2Click" />
                </menu>
              </splitButton>
              <toggleButton id="ToggleButton" label="Toggle Button" onAction="OnToggleButtonClick" />
              <dropDown id="DropDown" label="DropDown" onAction="OnDropDownSelected">
                <item id="DropDownItem1" label="Item 1" />
                <item id="DropDownItem2" label="Item 2" />
                <item id="DropDownItem3" label="Item 3" />
                <button id="button" label="Button..." />
              </dropDown>
              <comboBox id="ComboBox" label="Combo Box" onChange="OnComboBoxSelected">
                <item id="ComboBoxItem1" label="Item 1" />
                <item id="ComboBoxItem2" label="Item 2" />
                <item id="ComboBoxItem3" label="Item 3" />
              </comboBox>
              <separator id="separator" />
              <dynamicMenu id="DynamicMenu" label="Dynamic Menu" getContent="GetMenuContent" />
              <checkBox id="CheckBox" label="Check Box" onAction="OnCheckBoxToggled" />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>

**VBA-модуль SampleControls**

    Option Explicit

    Public Sub OnSplitButton1Click(rc As IRibbonControl)
        MsgBox "Split button 1 was clicked"
    End Sub

    Public Sub OnSplitButton2Click(rc As IRibbonControl)
        MsgBox "Split button 2 was clicked"
    End Sub

    Public Sub OnToggleButtonClick(rc As IRibbonControl, isButtonPressed As Boolean)
        MsgBox "Toggle button was toggled, button now is " & IIf(isButtonPressed, "pressed", "not pressed")
    End Sub

    Public Sub OnDropDownSelected(rc As IRibbonControl, selectedItemId As String, selectedItemIndex As Integer)
        MsgBox "DropDown was changed, selected item id is " & selectedItemId
    End Sub

    Public Sub OnComboBoxSelected(rc As IRibbonControl, comboBoxValue As String)
        MsgBox "Combo box was changed, value is " & comboBoxValue
    End Sub

    Public Sub GetMenuContent(rc As IRibbonControl, ByRef returnedVal)
        Dim xml As String

        xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & _
              "<button id=""but1"" imageMso=""Help"" label=""Help"" onAction=""OnHelpPressed""/>" & _
              "<button id=""but2"" imageMso=""FindDialog"" label=""Find"" onAction=""OnFindPressed""/>" & _
              "</menu>"

        returnedVal = xml
    End Sub

    Public Sub OnCheckBoxToggled(rc As IRibbonControl, isButtonChecked As Boolean)
        MsgBox "Check box was toggled, value is " & IIf(isButtonChecked, "checked", "not checked")
    End Sub

    Public Sub OnHelpPressed(rc As IRibbonControl)
        MsgBox "Help button pressed"
    End Sub

    Public Sub OnFindPressed(rc As IRibbonControl)
        MsgBox "Find button pressed"
    End Sub

Подробнее об элементах Custom UI: <a href="https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/d842006e-3187-4f66-a17d-0819a3cc94b5" rel="noopener" target="_blank">документация Microsoft</a>.

## Ещё бонус: горячие клавиши редактора VBE

**Навигация по редактору**

- <kbd>Ctrl+R</kbd> — перейти в окно проектов (Project Explorer)
- <kbd>F4</kbd> — перейти к свойствам
- <kbd>Ctrl+G</kbd> — перейти в окно Immediate
- <kbd>F7</kbd> — перейти в окно редактора кода
- <kbd>Ctrl+Tab</kbd> (<kbd>Ctrl+Shift+Tab</kbd>) — переключение между открытыми окнами модулей
- <kbd>Ctrl+F4</kbd> — закрыть текущий модуль
- <kbd>Alt+F11</kbd> — открыть редактор VBE

**Редактирование и навигация по коду**

- <kbd>Ctrl+Space</kbd> — автодополнение
- <kbd>Ctrl+J</kbd> — показать доступные варианты (IntelliSense)
- <kbd>Ctrl+I</kbd> — показать информацию о методе (аргументы)
- <kbd>Shift+F2</kbd> — перейти к определению переменной или метода

**Отладка**

- <kbd>F5</kbd> — запустить макрос или продолжить выполнение
- <kbd>F8</kbd> — шаг вперёд (с заходом во вложенные процедуры)
- <kbd>Shift+F8</kbd> — шаг вперёд (без захода вглубь)
- <kbd>F9</kbd> — установить или снять точку останова
- <kbd>Shift+F9</kbd> — добавить наблюдателя (quick watch)
- <kbd>Ctrl+L</kbd> — показать стек вызовов (элементы кликабельны)

О приёмах отладки подробнее в статье [«Приёмы отладки VBA-кода»](/ru/articles/debugging-vba-code-in-excel/).

Статья впервые опубликована на <a href="https://pikabu.ru/story/tutorial_sozdanie_vbanadstroyki_s_otdelnoy_vkladkoy_v_ms_excel_7160405" rel="noopener" target="_blank">Пикабу</a>.
