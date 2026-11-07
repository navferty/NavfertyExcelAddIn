{
  "title": "Creating a VBA add-in with its own ribbon tab in Excel",
  "description": "How to package VBA macros into an .xlam add-in with a custom ribbon tab: editing the workbook as a zip archive, customUI.xml, callbacks and installation.",
  "image": "https://navferty.net/images/articles/vba-addin-with-custom-ribbon-tab/1578783854145094499.png"
}

[← All articles](/en/articles/)

# Creating a VBA add-in with its own ribbon tab in Excel

How to package VBA macros into an .xlam add-in with a custom ribbon tab: editing the workbook as a zip archive, customUI.xml, callbacks and installation.

Many people who work in Excel a lot have built up a collection of VBA macros that make daily work easier. Some keep them in a separate workbook; others put them in a <a href="https://support.office.com/en-us/article/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566" rel="noopener" target="_blank">personal macro workbook</a> (personal.xlsb), which is available application-wide, and added buttons to the toolbar by hand. In the first case the collection is easy to share with colleagues (just send the file), but the workbook has to be opened every time you want to use it. In the second case the functions are available as soon as Excel starts, but passing them on to other users can be a problem.

This article describes how to make a VBA add-in with its own tab on the Excel ribbon (*Ribbon*), combining the advantages of both approaches. It is an xlam file that can be shared like an ordinary workbook. When a user opens it, a tab appears on the ribbon that gives access to the add-in functions:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578783854145094499.png" loading="lazy" width="700" height="108" alt="The new Sample Add-in tab on the Excel ribbon" />

If you wish, the add-in can be installed permanently so you do not have to open it each time.

You will need Excel, 7-Zip and Visual Studio Code. The last two are optional: any archiver that supports zip and any text editor that can save UTF-8 (if you want to use Cyrillic) will do. A basic understanding of XML helps too.

The code and the finished xlsm workbook are on GitHub: <a href="https://github.com/navferty/SampleExcelXlamAddin" rel="noopener" target="_blank">https://github.com/navferty/SampleExcelXlamAddin</a>

## Preparing the workbook

Open a new workbook and add a module with sample code:

    Option Explicit

    Public Sub SayHelloWorld()
        Dim firstCellValue As String

        firstCellValue = ThisWorkbook.Worksheets(1).Cells(1, 1).Value
        MsgBox "Hello, world! Cell A1 value is " & firstCellValue
    End Sub

Save the workbook as xlsm (macro-enabled workbook). Modern Office document formats (*xlsx, xlsm, docx, docm, pptx* and so on) are based on the OpenXML standard and are simply zip archives made of components (XML files for the sheets, images, binary content and so on) and relationship files (*.rels*) that define the structure of the document by linking to the components.

To see what is inside, you can rename *Sample.xlsm* to *Sample.zip* and open it with any archiver. Some archivers, such as 7-Zip, can open these files directly without renaming:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578771762187878642-en.jpg" loading="lazy" width="867" height="641" alt="The contents of an xlsm file in an archiver" />

To add a ribbon tab we need to create a new component, **customUI/customUI.xml**, and add a reference to it in the root relationships file, **\_rels/.rels**.

Extract the whole document into a separate folder (in 7-Zip: “Extract to "Sample\\”) and create a **customUI** folder in it:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/15787721721968742-en.jpg" loading="lazy" width="882" height="273" alt="The customUI folder in the unpacked workbook" />

Now add the component itself:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578772671166267580.webp" loading="lazy" width="700" height="427" alt="The customUI.xml file in the customUI folder" />

## The customUI.xml file

The text version of **customUI.xml**:

    <?xml version="1.0" encoding="utf-8"?>
    <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
      <ribbon>
        <tabs>
          <tab id="customTab" label="Sample Add-in" keytip="XT">
            <group id="Tools" label="Group 1">
              <button id="HelloWorldBtn" imageMso="BlackAndWhiteAutomatic" keytip="H" label="Hello, World!"
                      onAction="SayHelloWorld" screentip="Say hello to the world" size="large" />
              <button id="DuplicatesBtn" imageMso="SmartArtChangeColorsGallery" keytip="D"
                      label="Highlight duplicates" onAction="DuplicateColors"
                      screentip="Highlight repeated values in the selected range with paired colors" size="large" />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>

A short overview of the important elements and attributes.

- Identifier values (the *id* attribute) must be unique. For some elements (for example *tab*, when you declare a new tab rather than extend an existing one via *idMso*) you must also set *label*, otherwise Excel will not show the element.
- The *imageMso* attribute sets the element icon from the ones built into Office. The list of available values is long but easy to find with a search.
- The *button* element is an ordinary button. Depending on the *size* attribute it can be small or large, the full height of the tab.
- The *onAction* attribute names the procedure to run on the primary action of the element (for example a click on a button). The procedure must have a specific signature (the number and types of arguments), which you can find in the specification. For example, an *onAction* handler for a button takes one argument of type `IRibbonControl`.
- The procedure name must not match the module name (or you must prefix it with the module name explicitly: `Module1.SayHelloWorld`).
- The *keytip* attribute helps with keyboard navigation: the Alt key lets you move between Excel tabs.

A markup sample is available in the Microsoft documentation: <a href="https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-add-custom-ui-to-a-spreadsheet-document" rel="noopener" target="_blank">customUI sample</a>. If you want to study the standard in detail, there is a specification: <a href="https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43" rel="noopener" target="_blank">[MS-CUSTOMUI]: Custom UI XML Markup Specification</a>.

If the file contains non-ASCII characters (for example Cyrillic), save it as UTF-8, otherwise Excel may not show the new tab.

## The reference in the .rels file

For Excel to “see” the new component and understand what it is, we must reference it in the root relationships file, which is in the **\_rels** folder and is called **.rels**. It is XML too, and for convenience you can format it (in VS Code: Ctrl+Shift+P → Format Document).

Add a *Relationship* element with the path to customUI.xml, the component type and an identifier (it does not matter which, as long as it is unique among the other *Relationship* elements):

    <Relationship Id="rID4" Target="customUI/customUI.xml" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" />

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578772992128329742.webp" loading="lazy" width="700" height="427" alt="The .rels file with the added Relationship element" />

Now pack all the files back into an archive and restore the file name Sample.xlsm.

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878107816842521-en.jpg" loading="lazy" width="992" height="502" alt="Packing the files back into an archive" />

If everything was done correctly, a new tab appears on the ribbon when you open the workbook:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578781539166930736.png" loading="lazy" width="700" height="256" alt="The new tab on the ribbon after opening the workbook" />

## The “Wrong number of arguments” error

When you run the HelloWorld macro you may get this warning:

    Wrong number of arguments or invalid property assignment

This happens because a callback procedure must have a specific signature. A button click handler should look like this:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578781698127844347.webp" loading="lazy" width="700" height="332" alt="The signature of a button click handler" />

Add the `rc As IRibbonControl` argument and you will see the long-awaited greeting:

    Option Explicit

    Public Sub SayHelloWorld(rc As IRibbonControl)
        Dim firstCellValue As String

        firstCellValue = ThisWorkbook.Worksheets(1).Cells(1, 1).Value
        MsgBox "Hello, world! Cell A1 value is " & firstCellValue
    End Sub

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878173618865220.png" loading="lazy" width="211" height="133" alt="The greeting message box after clicking the button" />

The list of signatures is on the <a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)#how-can-i-determine-the-correct-signatures-for-each-callback-procedure" rel="noopener" target="_blank">Microsoft documentation site</a>.

## Saving as an add-in

All that remains is to save the workbook with macros as an add-in. In the VBE, select the `ThisWorkbook` object and set its `IsAddin` property to True.

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578782774119989254.webp" loading="lazy" width="700" height="445" alt="The IsAddin property of ThisWorkbook" />

Then save the workbook in the new format (Excel will offer this itself if you press Ctrl+S):

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878283811024502.webp" loading="lazy" width="700" height="394" alt="Saving the workbook as an add-in" />

The add-in can now be opened like an ordinary workbook. It is not shown in a separate window, but it is visible in the Project Explorer of the VBE:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878307317991772.webp" loading="lazy" width="700" height="497" alt="The add-in in the Project Explorer of the VBA editor" />

To make the add-in available right when Excel starts, install it from Excel Add-ins on the Developer tab...

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578783265155493627.webp" loading="lazy" width="700" height="479" alt="Installing the add-in from the Developer tab" />

...or from the Excel options, in the Add-ins section:

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/157878334215840375.webp" loading="lazy" width="700" height="574" alt="Installing the add-in from the Excel options" />

Done. The add-in can now be used with any document, and sharing it with colleagues is easy: just send them the add-in file.

## Bonus: other Custom UI controls

- **toggleButton** — a button with two states;
- **splitButton** — a composite of a button or toggleButton and a drop-down menu;
- **dropDown** — a drop-down list with a predefined set of items;
- **comboBox** — an input field that can hold a predefined set of items;
- **dynamicMenu** — a drop-down list whose items are defined dynamically in the getContent method;
- **checkBox** — a check box with two states.

<img src="/images/articles/vba-addin-with-custom-ribbon-tab/1578831453120134147.png" loading="lazy" width="700" height="108" alt="Examples of Custom UI controls on the ribbon" />

The markup and the VBA module code are below; you can also find them on <a href="https://github.com/navferty/SampleExcelXlamAddin" rel="noopener" target="_blank">GitHub</a>.

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

**SampleControls VBA module**

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

More about the Custom UI elements: <a href="https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/d842006e-3187-4f66-a17d-0819a3cc94b5" rel="noopener" target="_blank">Microsoft documentation</a>.

## One more bonus: VBE keyboard shortcuts

**Moving around the editor**

- <kbd>Ctrl+R</kbd> — go to the Project Explorer
- <kbd>F4</kbd> — go to the Properties window
- <kbd>Ctrl+G</kbd> — go to the Immediate window
- <kbd>F7</kbd> — go to the code editor window
- <kbd>Ctrl+Tab</kbd> (<kbd>Ctrl+Shift+Tab</kbd>) — switch between open module windows
- <kbd>Ctrl+F4</kbd> — close the current module
- <kbd>Alt+F11</kbd> — open the VBE

**Editing and code navigation**

- <kbd>Ctrl+Space</kbd> — auto-complete
- <kbd>Ctrl+J</kbd> — show the available options (IntelliSense)
- <kbd>Ctrl+I</kbd> — show information about a method (its arguments)
- <kbd>Shift+F2</kbd> — go to the definition of a variable or method

**Debugging**

- <kbd>F5</kbd> — run the macro or continue execution
- <kbd>F8</kbd> — step forward (into nested procedures)
- <kbd>Shift+F8</kbd> — step forward (without going inside)
- <kbd>F9</kbd> — set or remove a breakpoint
- <kbd>Shift+F9</kbd> — add a watch (quick watch)
- <kbd>Ctrl+L</kbd> — show the call stack (clickable)

There is more on debugging in [Debugging VBA code in Excel](/en/articles/debugging-vba-code-in-excel/).

Originally published in Russian on <a href="https://pikabu.ru/story/tutorial_sozdanie_vbanadstroyki_s_otdelnoy_vkladkoy_v_ms_excel_7160405" rel="noopener" target="_blank">Pikabu</a>.
