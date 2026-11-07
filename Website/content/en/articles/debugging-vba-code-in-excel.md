{
  "title": "Debugging VBA code in Excel",
  "description": "Breakpoints, step-into and step-over, the Immediate, Locals and Watches windows, conditional breakpoints and error trapping settings in the VBA editor.",
  "image": "https://navferty.net/images/articles/debugging-vba-code-in-excel/157842736017165403.png"
}

[← All articles](/en/articles/)

# Debugging VBA code in Excel

Breakpoints, step-into and step-over, the Immediate, Locals and Watches windows, conditional breakpoints and error trapping settings in the VBA editor.

Many people who work with Excel a lot use workbooks or add-ins with VBA macros. Often an existing solution needs to be extended or a bug fixed. Finding the place where the bug hides, or where a particular calculation happens, can take a long time, especially when a project has tens of thousands of lines of code (such macros do exist). Text search helps in some cases, for example by searching for the text of a message embedded in the code, but sometimes it is easier to run the macro and step through it.

This article describes a few ways to debug VBA macros in MS Office applications, using Excel as the example:

1.  Regular breakpoints, step-into, step-over and step-out.
2.  The Immediate window (a “REPL terminal”), the Locals window and the Watches window.
3.  Conditional breakpoints based on watch expressions.
4.  Error handling settings.

## 1. Breakpoints and stepping

The most obvious way to debug is to put a breakpoint at the start of the macro and then run it line by line, watching variable values and the state of the cells in the workbook. To set a breakpoint, press <span class="kbd">F9</span> and a marker appears on the line with the cursor. You can also click in the margin to the left of the code (the red dot).

<img src="/images/articles/debugging-vba-code-in-excel/1578427120196987602.webp" loading="lazy" width="700" height="214" alt="A breakpoint marker in the VBA editor" />

When execution reaches that line it pauses, and a yellow pointer shows the next statement to run:

<img src="/images/articles/debugging-vba-code-in-excel/157842736017165403.png" loading="lazy" width="424" height="47" alt="Execution paused, the yellow pointer marks the next statement" />

Not everyone knows that this pointer can be moved to any statement within the same procedure by dragging the yellow arrow with the mouse.

While execution is paused you can continue step by step with step-into, step-over and step-out:

- **step-into** (<span class="kbd">F8</span>) — one step forward. The line under the pointer is executed and the pointer moves to the next one. If the line calls another procedure, the pointer moves inside it.
- **step-over** (<span class="kbd">Shift+F8</span>) — the same as step-into, except that it does not go into nested calls; the called procedure runs in one go.
- **step-out** (<span class="kbd">Ctrl+Shift+F8</span>) — the current procedure runs to its end and execution pauses in the caller. Useful when you stepped into a long function by accident.

## 2. The Immediate, Locals and Watches windows

Now, how to find out the values of variables and other useful information. At the bottom of the Visual Basic Editor (VBE) there are three windows:

- Immediate — a terminal for running code and printing text;
- Locals — local variables;
- Watches — watch expressions.

<img src="/images/articles/debugging-vba-code-in-excel/157842820212856706.png" loading="lazy" width="700" height="132" alt="The Immediate, Locals and Watches windows at the bottom of the VBA editor" />

If you do not see these windows, enable them in the View menu.

<img src="/images/articles/debugging-vba-code-in-excel/157842841811498116.webp" loading="lazy" width="700" height="372" alt="The View menu with the debug window entries" />

Let us look at what these tools offer.

The **Immediate** window is a <a href="https://en.wikipedia.org/wiki/Read%E2%80%93eval%E2%80%93print_loop" rel="noopener" target="_blank">REPL</a> terminal: type an expression and press Enter to run it. To print the result in the same window, pass it to `Debug.Print` (you can also call that function from the macro code itself, and the output ends up in the Immediate window too). Inside the window you can use the short form, a question mark:

<img src="/images/articles/debugging-vba-code-in-excel/1578428865178198053.png" loading="lazy" width="523" height="270" alt="The Immediate window evaluating expressions and Debug.Print" />

Note that <span class="kbd">Ctrl+Space</span> and <span class="kbd">Ctrl+J</span> bring up IntelliSense here too, just like in the main code editor.

The **Locals** window shows the values of the variables available in the current context:

<img src="/images/articles/debugging-vba-code-in-excel/1578429103128329763.png" loading="lazy" width="527" height="256" alt="The Locals window showing variable values" />

For example, the current value of a loop counter or the contents of a collection. For `Collection` objects, expanding shows the values of the items, while for dictionaries (`Dictionary`) it shows their keys. So to see the value of a dictionary item you need to evaluate an expression such as:

    dict.Item("key1")

This can be done with a watch expression.

<img src="/images/articles/debugging-vba-code-in-excel/1578432168173270813.png" loading="lazy" width="512" height="94" alt="Adding a watch from the context menu" />

Select the expression and choose Add Watch... in the context menu (right mouse button), or type the expression by hand in the Add Watch dialog:

<img src="/images/articles/debugging-vba-code-in-excel/1578429559158522146.png" loading="lazy" width="399" height="298" alt="The Add Watch dialog" />

**Be careful:** evaluating `d.Item("key1")` adds an item to the dictionary if that key does not exist. Remember to delete the watch when you no longer need it, otherwise you may spend a long time later working out where the extra item in the dictionary came from.

## 3. Conditional breakpoints

Conditional breakpoints can be built from watches. Besides the standard Watch Expression option, the Add Watch dialog has two more: **Break When Value Is True** and **Break When Value Changes**. They pause the macro when the expression becomes true or when its value changes.

This is especially handy when debugging loops. If you know that a macro processing worksheet rows in a loop fails on, say, row 42, you can avoid stepping through the 41 previous iterations by hand. Add the expression `i=42` with the Break When Value Is True option, where `i` is the loop counter variable.

## 4. Error handling

Finally, a few words about error handling in VBA. In the VBE settings (Tools → Options, General tab) you can choose one of three options:

- Break on All Errors
- Break in Class Module
- Break on Unhandled Errors

<img src="/images/articles/debugging-vba-code-in-excel/1578430263138998521.png" loading="lazy" width="564" height="361" alt="Tools → Options → General: the Error Trapping settings" />

The last option is usually the default. It lets the program handle errors with `On Error GoTo` (show a message, clean up by closing a workbook or a connection) and `On Error Resume Next` (ignore the error and carry on). But when you run into a message like “Sorry, something went wrong” shown from an `On Error GoTo` handler, it is hard to tell where the error actually occurred. Break on All Errors helps here: it ignores the `On Error` directives and stops right on the line that raised the error. At that point the Locals window is useful too: the value of the loop counter, for example, tells you which worksheet row caused the error.

These settings are stored per user and are not saved in the VBA project.

I hope this is useful for people new to VBA as well as for experienced developers. If you work with Excel through an add-in and want to extend it, <a href="https://github.com/navferty/NavfertyExcelAddIn" rel="noopener" target="_blank">Navferty Excel Add-In</a> is open source under the MIT license, and suggestions are welcome in the GitHub issues.

Originally published in Russian on <a href="https://pikabu.ru/story/priyomyi_otladki_vbakoda_na_primere_ms_excel_7151540" rel="noopener" target="_blank">Pikabu</a>.
