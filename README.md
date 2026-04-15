# <p style="text-align:center;">SMTP Protocol Log Parser v2.1</p> 

<p style="text-align: center"><img src="https://www.cloudvision.com.tr/images/Logo/CV.png"></p>

[Cloud Vision](www.cloudvision.com.tr)


A Windows GUI tool for reading and analyzing Exchange Server SMTP Receive protocol logs. Built entirely in PowerShell 5.1 with no external dependencies - just run the script.

## Updated to v2.1 15.04.2026

Fixed file read issue while trying to open current active connector log file.

## Updated to v2 04.04.2026

Added SMTP Size sections and fixed Powershell direct calling. Also statistics are updated with more details .

## What it does

Exchange writes a lot of detail into its SMTP Receive protocol logs but they're just CSV files and not easy to read through raw. This tool loads one or more of those log files and lets you browse sessions, see what commands were exchanged, track emails through the session, and spot errors quickly.

A few things it handles that make it actually useful:

- Multiple emails per session (SMTP sessions that carry more than one message)
- Partial/incomplete sessions that didn't end cleanly
- Large log files - it parses in the background and shows progress so the UI stays responsive
- Multiple files at once - open a whole day's worth of logs together

## Requirements

- Windows PowerShell 5.1 (built into Windows Server 2016 and later, Windows 10 and later)
- No modules to install, no NuGet packages, nothing extra

## How to run

```
powershell.exe -File ProtocolLogParser.ps1   
```

or direct run with(in v2)

```
.\PrototolLogParser.ps1
```
Or right-click the file in Explorer and choose "Run with PowerShell". If you get a script execution policy error, run this first:

```
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## Log file format

The tool reads Exchange SMTP Receive connector protocol logs. These are typically found at:

```
C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive\
```

Files are plain CSV with a `#` comment header block at the top. The tool handles the header automatically.

## Views
<p style="text-align: center"><img src="https://raw.githubusercontent.com/burhanyurur/SMTPProtocolLogParser/refs/heads/main/main.png"></p>




**Sessions tab** - tree grouped by connector, then session, then individual emails. Clicking a session or email loads the raw log entries into the grid on the right.

**EHLO tab** - same sessions but grouped by the EHLO hostname the sending server announced. Handy for tracking down a specific sending host across multiple sessions.

**TLS tab** - shows TLS details per session (protocol version, cipher, certificate subject/issuer). Sessions without TLS are shown separately so you can spot unencrypted connections.

**Protocol View** - the main grid showing the raw log entries for whatever you selected in the tree. Click a row to see a decoded breakdown of the command or response in the panel below.

**Statistics** - charts for top senders, top recipients, error codes, and message volume by hour.

**Errors** - all error sessions grouped together with error codes and sample messages.

**Search** - filter sessions by sender IP, sender address, recipient, or session ID.

## Export
<p style="text-align: center"><img src="https://raw.githubusercontent.com/burhanyurur/SMTPProtocolLogParser/refs/heads/main/htmlexport.jpg"></p>


File > Export HTML Report generates a self-contained HTML file with the summary, charts, and session tables. No external CSS or JavaScript - the whole report is one file.

## Log file

The tool writes its own log (`ProtocolLogParser_YYYYMMDD.log`) next to the script for troubleshooting parse errors.

## Notes

- Session IDs in Exchange logs are unique per server but can repeat across different servers or after a service restart
- The tool marks sessions as "Incomplete" if there's no closing `-` event, which happens at log rotation boundaries or if the service was stopped mid-session
- Tested against Exchange 2016/2019/SE SMTP Receive logs
