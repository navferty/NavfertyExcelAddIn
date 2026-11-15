{
  "title": "Support and feedback",
  "description": "How to report a bug or request a feature for the Navferty Excel Add-In, where to find the logs and how to enable verbose logging.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[Home](/en/) › Support and feedback

# Support and feedback

Found a problem or missing a feature? Open an issue on GitHub. The **Info** button at the end of the **Navferty's Tools** ribbon tab opens a dialog with the same links.

<span id="bug"></span>

## Report a bug

<a href="https://github.com/navferty/NavfertyExcelAddIn/issues/new?template=bug_report.md" class="btn btn-primary btn-small" rel="noopener" target="_blank">Report a bug on GitHub</a>

Please include:

- the add-in version (copy it with the icon next to the version in the **Info** dialog) and your Excel version;
- the steps to reproduce the problem and what you expected to happen;
- a small sample workbook without sensitive data, if possible;
- the [log file](#logs), ideally recorded with [verbose logging](#verbose).

<span id="feature"></span>

## Request a feature

<a href="https://github.com/navferty/NavfertyExcelAddIn/issues/new?template=feature_request.md" class="btn btn-secondary btn-small" rel="noopener" target="_blank">Request a feature on GitHub</a>

Look through the <a href="https://github.com/navferty/NavfertyExcelAddIn/issues" rel="noopener" target="_blank">existing issues</a> first, and add your vote or comment if the idea is already there. Describe the problem you are trying to solve and a sample of your data rather than only the solution you have in mind.

<span id="logs"></span>

## Logs

The add-in writes a log file per day. To open the folder, click **Info** → **Open logs folder**, or go to:

`%LOCALAPPDATA%\Navferty\ExcelAddIn\logs`

Attach the latest `.log` file to your issue. Files larger than 5 MB are rotated, and files older than 14 days are deleted automatically.

<span id="verbose"></span>

## Verbose logging

By default only important events and errors are logged. To record detailed information:

1.  Click **Info** and tick **Verbose logging**.
2.  Reproduce the problem.
3.  Attach the log file to the issue.

Verbose logging is switched off again when Excel is restarted.

<div class="info-box">

### Privacy

Cell values, workbook and worksheet names and file paths are not written to the log: they are replaced with their length and a short hash. Error messages that come from Excel or other components are logged as they are, so please skim the file before attaching it to a public issue.

</div>
