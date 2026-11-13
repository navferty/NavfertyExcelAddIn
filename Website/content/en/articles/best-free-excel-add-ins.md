{
  "title": "Best free Excel add-ins",
  "description": "A short list of free add-ins that are worth installing, and what each is good at.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Best free Excel add-ins

A short list of free add-ins that are worth installing, and what each is good at.

Excel's ribbon has a lot of features, but many everyday chores still take too many steps. These free add-ins and built-in extras fill some of the gaps. Check each one's licence and privacy terms before using it with company data.

## Built into Excel (but switched off)

- **Analysis ToolPak** – regression, histograms, t-tests and other statistics. Enable in **File → Options → Add-ins → Manage: Excel Add-ins → Go**.
- **Solver** – optimisation under constraints, in the same dialog.
- **Power Query and Power Pivot** – import, transform and model data. Power Query is on the **Data** tab in current versions.

## From Microsoft

- **Microsoft Office Add-ins store** (**Insert → Get Add-ins**) has both free and paid items. The Microsoft Visio Data Visualizer, Power BI and a few chart tools are examples.
- **Power Query** for older versions (2010/2013) is a free separate download.

## Open-source and community

- **Navferty's Excel Add-In** – a free, open-source set of tools for working with cell contents: converting text to numbers, trimming spaces, replacing Cyrillic lookalike letters, transliteration, numbers to words (Russian, English, French), toggling case, unmerging cells with fill, highlighting duplicates, finding all error cells, validating values (including Russian tax IDs), exporting to SQLite and Markdown, creating a sample XML from an XSD and validating XML. It is a classic VSTO add-in for Windows desktop Excel. Source code and installer are on <a href="https://github.com/navferty/NavfertyExcelAddIn" rel="noopener" target="_blank">GitHub</a>.
- **Kutools**, **ASAP Utilities** and the **Ablebits** suite offer a lot more commands. They are free for limited use or for a trial period and then paid, so read the terms.
- **Excel-DNA** and **xlwings** are not add-ins to install but frameworks for developers: .NET functions in a worksheet, and Python with Excel.

## How to choose

- Does it do something you need at least weekly? A large toolbox that you use two buttons of is not a gain.
- Does it work offline and keep data on your machine? Check the privacy policy for anything that sends cells to a server.
- Can you see the source or at least trust the publisher?
- Does it fit your Excel (32 or 64 bit, version, Windows or Mac, desktop or web)? Many add-ins work only on Windows desktop.
- Can you uninstall it cleanly?

## Installing a third-party add-in

For installer-based add-ins, run the installer and restart Excel. Others come as `.xlam` or `.xll` files: put the file in a folder you trust, and add it in **File → Options → Add-ins → Go → Browse**. If Excel blocks it, right-click the file in Explorer → **Properties** and check **Unblock**. The [main page](/en/) has installation steps for Navferty's add-in.
