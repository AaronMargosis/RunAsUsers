# RunAsUsers

RunAsUsers runs arbitrary commands in the desktop sessions and security context of interactively logged-on users, including with the user's environment variables. If the user has a UAC "filtered token," RunAsUsers can optionally run the target program with the user's non-elevated or elevated token.

RunAsUsers doesn't assume there is at most one logged-on user, nor that the "Console" session is the only one that matters.

RunAsUsers can run programs hidden or with a visible interface that the user interacts with. It offers the option to redirect and capture the target programs' stdout and stderr to uniquely-named files in a designated directory.

RunAsUsers provides options to treat the input command line as a PowerShell command, and launches the appropriate PowerShell process.

RunAsUsers is a 32-bit executable. Its default behavior on 64-bit Windows is to _disable_ WOW64 file system redirection when processing the command line, so if a Windows executable is specified, it will be the one in System32 rather than in SysWOW64. RunAsUsers offers a command-line switch not to disable that redirection, favoring the execution of Windows 32-bit executables.

The file(s) specified in the command line must be in a location that is readable/executable by users. It is strongly recommended that the file(s) not also be modifiable by non-admin users.

RunAsUsers is a standalone command-line executable written in C++ with no DLL or framework dependencies.

<br>
<br>

---

## Command-line syntax:
<br>

> **RunAsUsers.exe [-s {first|active|all}] [-term** _n_ **|-wait** _n_ **|-wait inf] [-redirStd** _directory_ **[-merge]] [-e] [-hide|-min] [-p|-pb64|-pe] [-32] [-q] -c** _commandline_

<br>
Detailed description of command-line parameters:
<br><br>

| Parameter | Description |
| --- | --- |
| **-c** _commandline_ | Everything after the first **`-c`** becomes the command line to execute, with quotes preserved, etc. |
|||
|| **-s** specifies the desktop sessions in which to run the command line. |
| **-s first** | Run the command line only on the first active session found.|
| **-s active** | Run the command line in all active user sessions.|
| **-s all** | Run the command line in all logged-on sessions, whether active or disconnected. (This is the default.) |
| **-s _n_** | Run the command line in the session with ID _n_ (where _n_ is a positive decimal integer). |
|||
|| **-term** and **-wait** indicate whether and for how long to wait for processes to exit, and whether to terminate any that haven't completed. You must use one of these options to capture the processes' exit codes as well as to capture redirected stdout and stderr (see **-redirStd**).<br>If neither **-wait** nor **-term** is used, RunAsUsers does not wait for processes to exit, does not report their exit codes, and does not capture redirected stdout and stderr from those processes.|
| **-term** _n_ | Wait up to _n_ seconds for the process(es) to exit; terminate any that haven't exited.
| **-wait** _n_ | Wait up to _n_ seconds for the process(es) to exit; do not terminate any that are still running.
| **-wait inf** | Wait until all the launched processes have exited.|
|||
|**-redirStd** _directory_|Redirect the target processes' stdout and stderr to uniquely-named files in the named directory.<br>Use a hyphen **"-"** as the directory name to redirect the target processes' stdout/stderr to this process' stdout/stderr.<br>If a directory is specified, file names will incorporate session ID, process ID, timestamp, and whether it represents stdout or stderr output.<br>The **-redirStd** option is applicable only when using **-wait** or **-term** to monitor the target processes' output.<br>The named directory must already exist - RunAsUsers.exe will not create it.|
|**-merge**|When used with **-redirStd**, redirects each target process' stderr to its stdout.|
|||
|**-e**|Run the command line with the user's elevated permissions, if any. For example, if a user is a member of the Administrators group, **-e** will run the command line with full administrative rights; without **-e**, the command line will execute with the user's standard user rights.|
|**-hide**|Run the target process hidden (no UI). The default is to run it in its normal state (usually visible).|
|**-min**|Run the target process minimized to the taskbar.|
|||
||PowerShell options: treat _commandline_ as a PowerShell command and run it in a `powershell.exe` process. On 64-bit Windows it will run 64-bit `powershell.exe` unless the **-32** switch is also used.<br>`powershell.exe` will always be executed with the following options:<br>`-NoProfile -NoLogo -ExecutionPolicy Bypass`<br>and either `-Command` or `-EncodedCommand`.|
|**-p**|Pass _commandline_ to `powershell.exe` as-is with `-Command`.|
|**-pb64**|_commandline_ is already base64-encoded; pass it to `powershell.exe` as-is with `-EncodedCommand`.|
|**-pe**|Base64-encode the input _commandline_ and pass the result to `powershell.exe` with `-EncodedCommand`.|
|||
|**-32**|On 64-bit Windows, don't disable WOW64 file system redirection when executing _commandline_.<br>The default is to disable redirection and allow execution from the 64-bit System32 directory.|
|**-q**|Quiet mode: don't write detailed progress and diagnostic information to stdout.|
|||

<br>
<br>

---

## Usage examples:
<br>

---
### Run interactive programs on users' desktops

`RunAsUsers.exe -c notepad.exe C:\Windows\dxdiag.txt`<br>
On every logged-on user's desktop, opens 64-bit Notepad.exe with a text file in the Windows directory. 
<br>

`RunAsUsers.exe -hide -s active -c cmd.exe /d /c start https://www.wikipedia.org`<br>
On every active user's desktop, opens the user's default web browser to the Wikipedia home page.<br>
(The `-hide` switch keeps the `cmd.exe` console window from appearing, but does not hide the browser.)
<br><br>

---
### Empty users' recycle bins
`RunAsUsers.exe -p -hide -c Clear-RecycleBin -Force`<br>
Empty the Recycle Bin of every logged-on user, using the PowerShell cmdlet.
<br><br>

---
### Update User Configuration GPOs immediately for all logged-on users
`RunAsUsers.exe -wait 60 -redirStd - -hide -c gpupdate.exe /target:user /force`<br>
Run `gpupdate.exe` in every user session, capturing the gpupdate command-line output and exit codes.
(Changes to User Configuration won't otherwise take effect in existing sessions until GP updates on its own.)
<br><br>

---
### Log off one specific user session
`RunAsUsers.exe -s 2 -hide -c shutdown.exe /l /f`<br>
Force-log off the user session in session number 2.
<br><br>

---
### Log off all user sessions, without rebooting
`RunAsUsers.exe -s all -hide -c shutdown.exe /l /f`<br>
Force-log off each logged-on user session.
<br><br>

---
### Get user-specific file-system and gpresult information
`RunAsUsers.exe -p -hide -wait inf -redirStd C:\Logs -c '$env:USERPROFILE; dir -Force $env:USERPROFILE'`<br>
Get each logged-on user's userprofile directory and a list of that directory's contents.<br>
Each user's results will be stored in a uniquely-named file in the C:\Logs directory.
<br>

`RunAsUsers.exe -p -hide -wait inf -redirStd - -c '$env:USERPROFILE; dir -Force $env:USERPROFILE'`<br>
Get each logged-on user's userprofile directory and a list of that directory's contents.<br>
Each user's results are redirected to `RunAsUsers.exe`'s stdout.
<br>

`RunAsUsers.exe -hide -term 30 -c cmd.exe /d /c gpresult.exe /SCOPE USER /X C:\Temp\%USERNAME%.xml /F`<br>
Run GpResult.exe for each logged-on user to get that user's user configuration settings, written as an XML file with the username in the file name. RunAsUsers.exe terminates any GpResult.exe instance that has not completed after 30 seconds.<br>
`cmd.exe` is used to provide a way to get the value of each user's "USERPROFILE" environment variable.<br>
_Note_ that if `RunAsUsers.exe` is itself executed within a `cmd.exe` instance, each percent `%` character on the command line will need to be escaped with a caret `^` character so that it will be passed literally to the launched `cmd.exe` instance in each user session. For example:<br>
`cmd.exe /d /c RunAsUsers.exe -hide -term 30 -c cmd /d /c gpresult.exe /SCOPE USER /X C:\Temp2\^%USERNAME^%.xml /F`
<br><br>

---
### Determine whether users trust any unauthorized root certificates
`RunAsUsers.exe -pb64 -hide -term 30 -redirStd C:\Logs -c QwBvAG0AcABhAHIAZQAtAE8AYgBqAGUAYwB0ACAAKABnAGMAaQAgAEMAZQByAHQAOgBcAEwAbwBjAGEAbABNAGEAYwBoAGkAbgBlAFwAUgBvAG8AdAApACAAKABnAGMAaQAgAEMAZQByAHQAOgBcAEMAdQByAHIAZQBuAHQAVQBzAGUAcgBcAFIAbwBvAHQAKQAgAHwAIAA/AHsAIAAkAF8ALgBTAGkAZABlAEkAbgBkAGkAYwBhAHQAbwByACAALQBlAHEAIAAiAD0APgAiACAAfQAgAHwAIAAlAHsAIAAkAF8ALgBJAG4AcAB1AHQATwBiAGoAZQBjAHQAIAB9ACAAfAAgAGYAdAAgAFQAaAB1AG0AYgBwAHIAaQBuAHQALAAgAFMAdQBiAGoAZQBjAHQALAAgAEkAcwBzAHUAZQByACAALQBBAHUAdABvAFMAaQB6AGUA`<br>
This one-line command determines whether any logged-on user trusts one or more root certificates that aren't in the machine-wide store. Passed as a base64-encoded command to avoid challenges of quoting/escaping special characters in the command, and without having to save the commands to a script file (which would probably be easier, though). The original unencoded command is:<br>
`Compare-Object (gci Cert:\LocalMachine\Root) (gci Cert:\CurrentUser\Root) | ?{ $_.SideIndicator -eq "=>" } | %{ $_.InputObject } | ft Thumbprint, Subject, Issuer -AutoSize`<br> 
The resulting output in a user's stdout output file could look like this:<br>
```
Thumbprint                               Subject               Issuer
----------                               -------               ------
5B02EF61BADB350C7DCDAEED7BA5FE1A82C5532B CN=TotallyTrustworthy CN=TotallyTrustworthy
```
<br><br>

---
> [TODO - more coming here]
