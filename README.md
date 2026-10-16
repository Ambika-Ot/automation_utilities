# Desktop Tools

A small Windows desktop app that groups PowerPoint creation, Excel-to-XML
conversion, system metrics, and email attachments in one Tkinter interface.
The tools can also be run from the command line.

## Requirements

- Python 3.10 or newer
- The third-party packages used by the tools:
  - `python-pptx`
  - `lxml`
  - `openpyxl`
  - `psutil`
- Tkinter, which is included with most standard Windows Python installations

Install the packages from Command Prompt or PowerShell:

```powershell
py -m pip install python-pptx lxml openpyxl psutil
```

Keep the project files together in one folder. The PowerPoint tab also uses
`my_template.pptx` as its example template.

## Start the desktop app

From the project folder, run:

```powershell
py app.py
```

The app provides these tabs:

- **PowerPoint** — choose a template, select a theme or customize fonts and
  colors, replace `{{placeholders}}`, and add slides.
- **Excel to XML** — convert one workbook sheet or watch a folder for Excel
  files.
- **System Metrics** — record CPU, memory, disk, and network readings to CSV.
- **Send Mail** — select attachments, enter recipients and an optional subject,
  then open a console to send the email.

## PowerPoint

The PowerPoint tab creates a presentation from the selected `.pptx` template.
Text replacement values use the form `placeholder = value`; for example,
`company = Example Corp` replaces `{{company}}` in the presentation.

To use the command-line tool, create or edit a JSON config and run:

```powershell
py ppt_tool.py config_preset.json
```

Other example configs are `config_manual.json` and `config_mixed.json`. To list
available themes:

```powershell
py ppt_tool.py --themes
```

The config's `template` and `output` paths are resolved relative to the current
working directory. Update them if you run the command from a different folder.

## Excel to XML

The first row of the selected worksheet supplies XML field names. Blank rows
are skipped. Convert one workbook:

```powershell
py excel_to_xml.py input.xlsx
```

Specify an output path and worksheet:

```powershell
py excel_to_xml.py input.xlsx output.xml "Sheet2"
```

Watch a folder for `.xlsx` and `.xlsm` files:

```powershell
py excel_to_xml.py --watch "C:\path\to\incoming"
```

An optional second folder selects where generated XML is saved:

```powershell
py excel_to_xml.py --watch "C:\path\to\incoming" "C:\path\to\xml_output"
```

In watch mode, successfully converted workbooks are moved into a `processed`
subfolder. Failed workbooks are moved into `failed`. Press **Ctrl+C** to stop
the command-line watcher.

## System metrics

In the app, choose a CSV output file and a sampling interval of at least one
second, then start monitoring. The CSV contains a timestamp, CPU, RAM, disk,
and cumulative network traffic readings.

The standalone metrics script writes six readings to `metrics.csv` in the
current working directory:

```powershell
py metric_data.py
```

## Send mail

The Send Mail tab opens a console and runs `cmd_mail.py` with the selected
attachments and recipients. Enter email addresses separated by commas or
semicolons. The script prompts for the sender account and password in the
console.

For Gmail, use a Google App Password if required by your account; do not use
your regular account password. The mail script also supports `MAIL_USER` and
`MAIL_PASS` environment variables, but avoid storing credentials in source
files or committing them to version control.

The command-line script can also be run directly:

```powershell
py cmd_mail.py --to recipient@example.com --subject "Report" --files "C:\path\to\report.pdf"
```

## Project files

| File | Purpose |
| --- | --- |
| `app.py` | Tkinter desktop application and its tabs |
| `ppt_tool.py` | PowerPoint generation and theme customization |
| `excel_to_xml.py` | Excel-to-XML conversion and folder watching |
| `metric_data.py` | System metric sampling and CSV output |
| `cmd_mail.py` | Email sending with file attachments |
| `config_*.json` | Example PowerPoint configurations |
| `my_template.pptx` | Example PowerPoint template |

