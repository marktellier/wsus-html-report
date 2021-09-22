# WSUS HTML Report
![servers](/assets/html-report.png)

This is another WSUS report generated as a dynamic html report.

This script requires the PoshWSUS module from PowerShell Gallery and must be installed on your WSUS server, it's simple as this:

```powershell
Install-Module PoshWSUS
```

At minimum, modify the following variables

| Variable      | Description                                |
| ------------- | ------------------------------------------ |
| $smtpServer   | IP Address of your SMTP Server             |
| $smtpFrom     | From email address                         |
| $smtpto       | Email recipient(s)                         |
| $reportOutput | Path location, be sure it's a valid folder |

The secret sauce to this report is really the html styles that are defined, play with these settings to get the report to display to your liking.

After the report is run, it's saved to C:\Reports folder and opens with your default web browser. The file is both embedded and attached to an email so that you can see the differences.

For my preferences, the embedded spreadsheet is more convenient but looses highlighting and creates some double horizontal lines. The attached file has more flash but has to be saved or opened and less useful on a mobile device.

Have fun with the script.
