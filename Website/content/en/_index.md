{
  "title": "Navferty's Excel Add-In",
  "description": "A free, open-source Excel add-in with commands for cleaning text, converting numbers, finding errors and duplicates, validating values and XML, and exporting to SQLite.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}


<section class="hero">
<h1 class="zag">Navferty's Excel Add-In</h1>
<p class="subtitle">A free Excel add-in that enhances productivity with tools for data manipulation, formatting, and validation.</p>
<ul class="facts"><li>Free</li><li>Open source (MIT)</li><li>For desktop Excel on Windows</li></ul>
<div class="hero-actions">
<a class="btn btn-primary" href="https://deploy.navferty.net/setup.zip"><svg aria-hidden="true" fill="none" height="20" role="img" viewbox="0 0 20 20" width="20" xmlns="http://www.w3.org/2000/svg"><path d="M10 13L5 8h3V3h4v5h3l-5 5z" fill="currentColor"></path><path d="M3 15h14v2H3v-2z" fill="currentColor"></path></svg>Download installer</a>
<a class="btn btn-secondary" href="#commands">Browse commands</a>
</div>
<img alt="Navferty's Tools Ribbon in Microsoft Excel" class="hero-image" height="332" src="/images/NavfertyToolsRibbonEn.png" width="1687"/>
</section>
<section id="guides">
<h2>Guides for everyday Excel problems</h2>
<p>Most of these start with the problem and show the built-in way to solve it. The add-in is mentioned where it saves some clicks.</p>
<div class="card-grid">
<a class="card" href="/en/articles/find-all-errors-in-excel/"><strong>How to find all errors in Excel (#VALUE!, #REF!, #DIV/0!)</strong><span>Ways to locate every formula error in a sheet, and what each error usually means.</span></a>
<a class="card" href="/en/articles/text-numbers-to-numeric-values/"><strong>Converting text numbers to numeric values</strong><span>Why numbers become text in Excel and several ways to convert them back.</span></a>
<a class="card" href="/en/articles/remove-extra-spaces-in-excel/"><strong>Removing extra spaces in Excel</strong><span>TRIM, non-breaking spaces and line breaks: how to clean text cells properly.</span></a>
<a class="card" href="/en/articles/highlight-duplicate-values-in-excel/"><strong>Highlighting duplicate values without formulas</strong><span>Conditional formatting, Remove Duplicates and what to do when you need to see which rows match.</span></a>
<a class="card" href="/en/articles/vlookup-in-excel/"><strong>VLOOKUP in Excel: syntax, examples and common errors</strong><span>Syntax, a worked example with two sheets, the usual reasons for #N/A, and XLOOKUP or INDEX/MATCH as alternatives.</span></a>
<a class="card" href="/en/articles/excel-keyboard-shortcuts/"><strong>Excel keyboard shortcuts that save the most time</strong><span>Shortcuts for editing, moving, selecting, filling and pasting, grouped by task.</span></a>
</div>
<p class="more-link"><a href="/en/articles/">All guides →</a></p>
</section>
<section id="commands">
<h2>Commands</h2>
<p>Everything on the Navferty's Tools ribbon tab. Open a command to see what it does with before/after screenshots.</p>
<h3>Clean up text and numbers</h3>
<div class="cmd-grid">
<a class="cmd" href="/en/features/parse-numerics/"><span class="cmd-thumb"><img alt="Example of values converted from text to numbers in Excel." height="426" loading="lazy" src="/images/ParseNumerics2.png" width="328"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/parseNumerics.png" width="64"/><span class="cmd-text"><strong>Parse Numerics</strong><span>Turns numbers stored as text into real numbers.</span></span></span></a>
<a class="cmd" href="/en/features/trim-spaces/"><span class="cmd-thumb"><img alt="Example of text cells after extra spaces and line breaks are removed." height="367" loading="lazy" src="/images/TrimSpaces2.png" width="419"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/trimSpaces.png" width="64"/><span class="cmd-text"><strong>Trim Spaces</strong><span>Removes extra spaces and line breaks from text cells.</span></span></span></a>
<a class="cmd" href="/en/features/toggle-case/"><span class="cmd-thumb"><img alt="Example of text with its letter case changed in Excel." height="206" loading="lazy" src="/images/ToggleCase2.png" width="161"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/toggleCase.png" width="64"/><span class="cmd-text"><strong>Toggle Case</strong><span>Cycles text between Abc, abc and ABC.</span></span></span></a>
<a class="cmd" href="/en/features/unmerge-cells/"><span class="cmd-thumb"><img alt="Example of merged cells unmerged and filled with their original value." height="199" loading="lazy" src="/images/Unmerge2.png" width="550"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/unmergeCells.png" width="64"/><span class="cmd-text"><strong>Unmerge Cells</strong><span>Unmerges cells and fills each one with the original value.</span></span></span></a>
<a class="cmd" href="/en/features/repair-conditional-format/"><span class="cmd-thumb"><img alt="Example of conditional formatting filled across selected rows using a source row." height="646" loading="lazy" src="/images/ConditionalFormatting2.png" width="716"/></span><span class="cmd-body"><span class="cmd-text"><strong>Conditional Formatting</strong><span>Fills formatting across selected rows using a source row.</span></span></span></a>
</div>
<h3>Letters and number formats</h3>
<div class="cmd-grid">
<a class="cmd" href="/en/features/replace/"><span class="cmd-thumb"><img alt="Example of Cyrillic text transliterated into Latin characters." height="206" loading="lazy" src="/images/Transliterate2.png" width="454"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/replace.png" width="64"/><span class="cmd-text"><strong>Replace</strong><span>Transliterates Cyrillic or swaps look-alike letters for Latin ones.</span></span></span></a>
<a class="cmd" href="/en/features/stringify-numerics/"><span class="cmd-thumb"><img alt="Example of numeric values written out as words." height="307" loading="lazy" src="/images/Stringify2.png" width="883"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/stringifyNumerics.png" width="64"/><span class="cmd-text"><strong>Stringify Numerics</strong><span>Writes numbers in words in Russian, English or French.</span></span></span></a>
</div>
<h3>Find problems in data</h3>
<div class="cmd-grid">
<a class="cmd" href="/en/features/find-errors/"><span class="cmd-thumb"><img alt="Example of formula error cells found in a selected range." height="611" loading="lazy" src="/images/FindErrorValues.png" width="968"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/findErrors.png" width="64"/><span class="cmd-text"><strong>Find All Cells Containing Errors</strong><span>Lists every formula error in the selection and jumps to it.</span></span></span></a>
<a class="cmd" href="/en/features/highlight-duplications/"><span class="cmd-thumb"><img alt="Example of duplicate values highlighted with a different colour for each group." height="306" loading="lazy" src="/images/Duplicates.png" width="858"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/highlightDuplicates.png" width="64"/><span class="cmd-text"><strong>Highlight Duplications</strong><span>Colours repeated values, one colour per group.</span></span></span></a>
<a class="cmd" href="/en/features/validate-values/"><span class="cmd-thumb"><img alt="Example of selected spreadsheet values checked by the validation command." height="607" loading="lazy" src="/images/Validate.png" width="1130"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/validation.png" width="64"/><span class="cmd-text"><strong>Validate Values</strong><span>Checks values as numbers, dates, TIN or with an XML-text character filter.</span></span></span></a>
</div>
<h3>Export, XML and protection</h3>
<div class="cmd-grid">
<a class="cmd" href="/en/features/copy-as-markdown/"><span class="cmd-thumb"><img alt="Example of selected Excel cells prepared as a Markdown table." height="306" loading="lazy" src="/images/ExportToMarkdown.png" width="858"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/markdown.png" width="64"/><span class="cmd-text"><strong>Copy as Markdown</strong><span>Copies the selected cells as a Markdown table.</span></span></span></a>
<a class="cmd" href="/en/features/export-to-sqlite/"><span class="cmd-thumb"><img alt="Example of worksheet data prepared for export to a SQLite database." height="900" loading="lazy" src="/images/ExportToSqlite2.png" width="1412"/></span><span class="cmd-body"><img alt="" class="icons" height="64" loading="lazy" src="/images/icons/validation.png" width="64"/><span class="cmd-text"><strong>Export to SQLite</strong><span>Exports worksheet data to a SQLite database.</span></span></span></a>
<a class="cmd" href="/en/features/create-sample-xml/"><span class="cmd-thumb"><img alt="Example XML document generated from an XSD schema." height="441" loading="lazy" src="/images/CreateSampleXml1.png" width="497"/></span><span class="cmd-body"><span class="cmd-text"><strong>Create Sample XML based on XSD</strong><span>Builds a sample XML file from an XSD schema.</span></span></span></a>
<a class="cmd" href="/en/features/validate-xml/"><span class="cmd-thumb"><img alt="Example of XML validation results against an XSD schema." height="257" loading="lazy" src="/images/ValidateXml1.png" width="883"/></span><span class="cmd-body"><span class="cmd-text"><strong>Validate XML with XSD</strong><span>Checks an XML file against an XSD and reports errors in a sheet.</span></span></span></a>
<a class="cmd" href="/en/features/unprotect-workbook/"><span class="cmd-thumb"><img alt="Example of workbook protection options in the add-in." height="601" loading="lazy" src="/images/UnprotectWorkbook1.png" width="663"/></span><span class="cmd-body"><span class="cmd-text"><strong>Unprotect Workbook</strong><span>Removes sheet and workbook-structure protection; attempts to unlock VBA.</span></span></span></a>
<a class="cmd" href="/en/features/worksheets-protection/"><span class="cmd-thumb"><img alt="Example of protection controls for multiple worksheets." height="573" loading="lazy" src="/images/WorksheetsProtection1.png" width="818"/></span><span class="cmd-body"><span class="cmd-text"><strong>Worksheets Protection</strong><span>Protects or unprotects many worksheets at once.</span></span></span></a>
<a class="cmd" href="/en/features/undo-last-action/"><span class="cmd-thumb cmd-thumb-icon"><img alt="" height="64" loading="lazy" src="/images/icons/undo.png" width="64"/></span><span class="cmd-body"><span class="cmd-text"><strong>Undo Last Action</strong><span>Reverts the most recent supported value transformation.</span></span></span></a>
</div>
<h3>More tools</h3>
<div class="cmd-grid">
<a class="cmd" href="/en/features/toggle-sheet-tabs/"><span class="cmd-thumb"><img alt="Example of worksheet tabs shown or hidden in Excel." height="621" loading="lazy" src="/images/ToggleSheetTabs2.png" width="1245"/></span><span class="cmd-body"><span class="cmd-text"><strong>Sheet Tabs</strong><span>Shows or hides the worksheet tabs.</span></span></span></a>
<a class="cmd" href="/en/features/switch-reference-style/"><span class="cmd-thumb"><img alt="Example of formulas displayed with A1 or R1C1 cell references." height="453" loading="lazy" src="/images/SwitchReferenceStyle2.png" width="1288"/></span><span class="cmd-body"><span class="cmd-text"><strong>Reference Style (A1/R1C1)</strong><span>Switches formula references between A1 and R1C1 style.</span></span></span></a>
<a class="cmd" href="/en/features/exchange-rates/"><span class="cmd-thumb"><img alt="Example of an exchange rate inserted into an Excel cell." height="173" loading="lazy" src="/images/ExchangeRates2.png" width="508"/></span><span class="cmd-body"><span class="cmd-text"><strong>Exchange rates</strong><span>Picks a currency rate from an online source and pastes it into the cell.</span></span></span></a>
</div>
</section>
<section id="install">
<h2>Installation</h2>
<div class="installation-steps">
<div class="step"><div class="step-number">1</div><div class="step-content"><h4>Download and unpack the installer archive</h4><p>Get the latest version from our server:</p>
<a class="btn btn-primary btn-small" href="https://deploy.navferty.net/setup.zip">Download setup.zip</a></div></div>
<div class="step"><div class="step-number">2</div><div class="step-content"><h4>Run the installation</h4><p>Double-click the downloaded <code>setup.exe</code> file and follow the installation wizard.</p></div></div>
<div class="step"><div class="step-number">3</div><div class="step-content"><h4>Restart Excel</h4><p>If Excel was open during installation, restart it. You'll see the new <strong>Navferty's Tools</strong> tab in the ribbon.</p></div></div>
</div>
<p class="more-link"><a href="/en/install/">Full instructions, certificate and offline install →</a></p>
</section>
<section id="support">
<h2>Support and feedback</h2>
<p>Found a bug or need a new command? <a href="/en/support/#bug">Report a bug</a> or <a href="/en/support/#feature">request a feature</a> on GitHub, and see how to <a href="/en/support/#logs">find the logs</a> to attach.</p>
<p class="more-link"><a href="/en/support/">Support and feedback →</a></p>
</section>
