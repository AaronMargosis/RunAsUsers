// RunAsUsers.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <Windows.h>
#include <UserEnv.h>
#pragma comment(lib, "UserEnv.lib")
#include <WtsApi32.h>
#pragma comment(lib, "wtsapi32.lib")
#include <io.h>
#include <fcntl.h>
#include <iostream>
#include <sstream>
#include <vector>
#include "ProcessManager.h"
#include "SecUtils.h"
#include "SysErrorMessage.h"
#include "UtilityFunctions.h"
#include "Wow64FsRedirection.h"
#include "HEX.h"
#include "DbgOut.h"
#include "RedirManager.h"
#include "StringUtils.h"
#include "WhoAmI.h"
#include "Token.h"

// Considered adding -o outfile and -o2 errfile command line options, but this process writes to stdout/stderr through 
// std::wcout/std::wcerr and through WriteFile (see RedirManager.cpp). Unless/until I come up with a way to redirect
// both of those to the same file without stepping on each other's output, it's not going to be feasible to redirect 
// it from within the process.

// Considered providing options so that PowerShell didn't always include "-ExecutionPolicy Bypass" but decided not worth
// it. Orgs that want to maintain tighter control over PS execution policy will do so using group policy, which takes
// precedence over -ExecutionPolicy.

static const wchar_t* const szPowerShellCmd = L"powershell.exe -NoProfile -NoLogo -ExecutionPolicy Bypass";

/// <summary>
/// Write command-line syntax to stderr (with optional error information) and then exit
/// </summary>
/// <param name="argv0">The program's argv[0] value</param>
/// <param name="szError">Optional: caller-supplied error text</param>
/// <param name="szBadParam">Optional: additional caller-supplied error text (invalid parameter)</param>
void Usage(const wchar_t* argv0, const wchar_t* szError = nullptr, const wchar_t* szBadParam = nullptr)
{
    std::wstring sExe = GetFileNameFromFilePath(argv0);
    if (szError)
    {
        std::wcerr << szError;
        if (szBadParam) 
            std::wcerr << L": " << szBadParam;
        std::wcerr << std::endl;
    }
    std::wcerr
        << std::endl
        << L"Usage:" << std::endl
        << std::endl
        << L"  " << sExe << L" [-s {first|active|all|n}] [-wait n | -wait inf | -term n] [-redirStd directory [-merge]] [-e] [-hide|-min] [-p|-pb64|-pe] [-32] [-q] -c commandline" << std::endl
        << std::endl
        << L"    -c commandline" << std::endl
        << L"      Everything after the first -c becomes the command line to execute, with quotes preserved, etc." << std::endl
        << std::endl
        << L"    -s : desktop session(s) in which to execute:" << std::endl
        << L"        -s first" << std::endl
        << L"          Run the command line only on the first active session found." << std::endl
        << L"        -s active" << std::endl
        << L"          Run the command line in all active user sessions." << std::endl
        << L"        -s all" << std::endl
        << L"          Run the command line in all logged-on sessions, whether active or disconnected." << std::endl
        << L"        -s n" << std::endl
        << L"          Run the command line in the session with ID \"n\" (where \"n\" is a positive decimal integer)." << std::endl
        << L"      If -s is not used, \"all\" is the default." << std::endl
        << std::endl
        << L"    -wait, -term : whether and how long to wait for process(es) to exit, and report exit code:" << std::endl
        << L"        -term n" << std::endl
        << L"          Wait up to n seconds for the process(es) to exit; terminate it if it hasn't exited." << std::endl
        << L"        -wait n" << std::endl
        << L"          Wait up to n seconds for the process(es) to exit; do not terminate if it hasn't." << std::endl
        << L"        -wait inf" << std::endl
        << L"          Wait until the process(es) exit and report exit code." << std::endl
        << L"      If neither -wait or -term is used, " << sExe << L" does not wait for processes to exit," << std::endl
        << L"      and does not report exit codes." << std::endl
        << std::endl
        << L"    -redirStd directory" << std::endl
        << L"      Redirect the target processes' stdout and stderr to uniquely-named files in the named directory." << std::endl
        << L"      Use \"-\" as the directory name to output target processes' stdout and stderr through this process' stdout/stderr." << std::endl
        << L"      Add -merge to redirect the target processes' stderr to its stdout." << std::endl
        << L"      -redirStd is applicable only when using -wait or -term to monitor the target process' output." << std::endl
        << std::endl
        << L"    -e" << std::endl
        << L"      Run the command line with the user's elevated permissions, if any." << std::endl
        << std::endl
        << L"    -hide" << std::endl
        << L"      Run the target process hidden (no UI). The default is to run it in its normal state (usually visible)." << std::endl
        << L"    -min" << std::endl
        << L"      Run the target process minimized to the taskbar." << std::endl
        << std::endl
        << L"    -p, -pb64, -pe : runs the command line in the target sessions as a PowerShell command:" << std::endl
        << L"        -p" << std::endl
        << L"          Pass the \"commandline\" argument to powershell.exe as-is with -Command." << std::endl
        << L"        -pb64" << std::endl
        << L"          \"commandline\" is already base-64 encoded; pass it to powershell.exe as-is with -EncodedCommand." << std::endl
        << L"        -pe" << std::endl
        << L"          Base64-encode the input \"commandline\" and pass the result to powershell.exe with -EncodedCommand." << std::endl
        << L"      With each of the above, PowerShell.exe is invoked with these command-line switches:" << std::endl
        << L"      " << szPowerShellCmd << std::endl
        << std::endl
        << L"    -32" << std::endl
        << L"      Don't disable WOW64 file system redirection when executing the command line." << std::endl
        << L"      (Default is to disable redirection and allow execution from the 64-bit System32 directory.)" << std::endl
        << std::endl
        << L"    -q" << std::endl
        << L"      Quiet - don't write detailed progress/diagnostic information to stdout." << std::endl
        << std::endl;
    exit(-1);
}



// ----------------------------------------------------------------------------------------------------------------------------------
/// <summary>
/// Get the target command line to execute: everything past the first "-c" on the command line.
/// </summary>
/// <param name="sCommandLine">Output: the target command line to execute</param>
/// <returns>true if command line found, false otherwise</returns>
static bool GetTargetCommandLine(std::wstring& sCommandLine)
{
    sCommandLine.clear();

    const wchar_t* szTargetCmdline = GetCommandLineW();
    const wchar_t* szDashC = wcsstr(szTargetCmdline, L" -c ");
    if (nullptr != szDashC)
    {
        szTargetCmdline = szDashC + 4;
        if (*szTargetCmdline)
        {
            sCommandLine = szTargetCmdline;
            return true;
        }
    }
    return false;
}

/// <summary>
/// Return the directory that should be used as the current directory for target processes.
/// Current implementation uses the system directory (typically C:\Windows\System32), which will work
/// whether target processes are 32- or 64-bit.
/// </summary>
/// <returns></returns>
static const std::wstring& TargetCurrentDirectory()
{
    // Get the value on first use, then reuse that value.
    static std::wstring sTargetCurrentDirectory;
    if (0 == sTargetCurrentDirectory.length())
    {
        wchar_t szBuffer[MAX_PATH] = { 0 };
        if (GetSystemDirectoryW(szBuffer, MAX_PATH) > 0)
        {
            sTargetCurrentDirectory = szBuffer;
        }
    }
    return sTargetCurrentDirectory;
}

/// <summary>
/// Enum for the -s option (which session or sessions to execute processes in)
/// </summary>
enum class WhichSessions_t {
    firstActive,
    allActive,
    allLoggedOn,
    oneSessionId
};

/// <summary>
/// Convert the WhichSesssions_t enum to corresponding string.
/// </summary>
inline std::wstring WhichSessionsToWSZ(WhichSessions_t whichSessions, DWORD nSessionId)
{
    switch (whichSessions)
    {
    case WhichSessions_t::firstActive:
        return L"First active";
    case WhichSessions_t::allActive:
        return L"All active";
    case WhichSessions_t::allLoggedOn:
        return L"All logged on";
    case WhichSessions_t::oneSessionId:
    {
        std::wstringstream str;
        str << L"Session # " << nSessionId;
        return str.str();
    }
    default:
        return L"UNEXPECTED/UNDEFINED";
        break;
    }
}

// Commented-out wmainImpl code for stress- and leak-testing
//int wmainImpl(int argc, wchar_t** argv);

int wmain(int argc, wchar_t** argv)
//{
//    std::wstring s;
//    std::wcout << L"READY TO START HEAVY TEST. PRESS KEYS AND ENTER: " << std::endl;
//    std::wcin >> s;
//
//    for (size_t ix = 0; ix < 50; ++ix)
//    {
//        std::wcout << L"BEGIN ITERATION " << ix << std::endl;
//        //std::wcout << L"argc = " << argc << L"; ";
//        //for (int ixArg = 0; ixArg < argc; ++ixArg)
//        //    std::wcout << argv[ixArg] << L", ";
//        //std::wcout << std::endl;
//        wmainImpl(argc, argv);
//        std::wcout << L"ENDING ITERATION " << ix << std::endl;
//    }
//
//    std::wcout << L"DONE WITH HEAVY TEST. PRESS KEYS AND ENTER: " << std::endl;
//    std::wcin >> s;
//    return 0;
//}
//int wmainImpl(int argc, wchar_t** argv)
{
    // Default is no debug output; can be changed with hidden options
    dbgOut.WriteToDebugStream(false);

    // Set output mode to UTF8.
    if (_setmode(_fileno(stdout), _O_U8TEXT) == -1 || _setmode(_fileno(stderr), _O_U8TEXT) == -1)
    {
        std::wcerr << L"Unable to set stdout and/or stderr modes to UTF8." << std::endl;
    }

    if (argc < 3)
        Usage(argv[0]);

    std::wstring 
        sOriginalCommandLine, 
        sActualCommandLine, 
        sRedirStdDirectory,
        sDbgLogFname;
    bool
        bQuiet = false,
        bPowerShell = false, bIsBase64 = false, bEncode = false,
        bRedirStd = false,
        bMergeStd = false,
        bTryElevated = false,
        bHidden = false,
        bMinimized = false,
        bTerminate = false,
        bWow64FileSystemRedir = false,
        bDebug = false, bDebugF = false;
    DWORD 
        dwWait = 0, 
        nSessionId = 0,
        nTargetedSessions = 0;
    WhichSessions_t whichSessions = WhichSessions_t::allLoggedOn;

    DWORD dwLastErr = 0;

    // Get the target command line to execute (everything past the first "-c")
    if (!GetTargetCommandLine(sOriginalCommandLine))
        Usage(argv[0], L"Missing command line");

    // Parse command-line options
    int ixArg = 1;
    bool bDoneParsing = false;
    while (ixArg < argc && !bDoneParsing)
    {
        // Ignore everything past the first -c -- already picked that up in GetTargetCommandLine
        if (0 == wcscmp(L"-c", argv[ixArg]))
        {
            bDoneParsing = true;
        }
        else if (0 == wcscmp(L"-p", argv[ixArg]))
        {
            // Can use at most only one PowerShell option
            if (bPowerShell)
                Usage(argv[0], L"PowerShell option already specified");
            // Standard PowerShell execution
            bPowerShell = true;
        }
        else if (0 == wcscmp(L"-pb64", argv[ixArg]))
        {
            // Can use at most only one PowerShell option
            if (bPowerShell)
                Usage(argv[0], L"PowerShell option already specified");
            // Input PowerShell command is base64-encoded
            bPowerShell = true;
            bIsBase64 = true;
        }
        else if (0 == wcscmp(L"-pe", argv[ixArg]))
        {
            // Can use at most only one PowerShell option
            if (bPowerShell)
                Usage(argv[0], L"PowerShell option already specified");
            // Input PowerShell command should be base64-encoded before passing to PowerShell
            bPowerShell = true;
            bEncode = true;
        }
        else if (0 == wcscmp(L"-redirStd", argv[ixArg]))
        {
            // Redirect target process' stdout/stderr to uniquely-named files in named directory
            if (++ixArg >= argc)
                Usage(L"Missing arg for -redirStd", argv[0]);
            sRedirStdDirectory = argv[ixArg];
            bRedirStd = true;
        }
        else if (0 == wcscmp(L"-merge", argv[ixArg]))
        {
            bMergeStd = true;
        }
        else if (0 == wcscmp(L"-e", argv[ixArg]))
        {
            // Run the target executable using each user's elevated token, if possible.
            bTryElevated = true;
        }
        else if (0 == wcscmp(L"-hide", argv[ixArg]))
        {
            // Run the target executable hidden (no visible UI)
            // Note: if both specified, hidden takes precedence over minimized
            bHidden = true;
        }
        else if (0 == wcscmp(L"-min", argv[ixArg]))
        {
            // Run the target executable minimized
            // Note: if both specified, hidden takes precedence over minimized
            bMinimized = true;
        }
        else if (0 == wcscmp(L"-term", argv[ixArg]))
        {
            // Terminate the target executable if it hasn't completed within specified wait period
            // 
            // -term/-wait can be used once at most
            if (0 != dwWait)
                Usage(argv[0], L"-term/-wait can be specified once at most");

            if (++ixArg >= argc)
                Usage(argv[0], L"Missing arg for -term");
            if (1 != swscanf_s(argv[ixArg], L"%lu", &dwWait) || 0 == dwWait)
                Usage(argv[0], L"Invalid arg for -term", argv[ixArg]);
            bTerminate = true;
        }
        else if (0 == wcscmp(L"-wait", argv[ixArg]))
        {
            // Wait for target executable for specified amount of time to report exit code
            // 
            // -term/-wait can be used once at most
            if (0 != dwWait)
                Usage(argv[0], L"-term/-wait can be specified once at most");

            if (++ixArg >= argc)
                Usage(argv[0], L"Missing arg for -wait");
            if (0 == wcscmp(L"inf", argv[ixArg]))
                dwWait = INFINITE;
            else if (1 != swscanf_s(argv[ixArg], L"%lu", &dwWait) || 0 == dwWait)
                Usage(argv[0], L"Invalid arg for -wait", argv[ixArg]);
        }
        else if (0 == wcscmp(L"-s", argv[ixArg]))
        {
            // Which session(s) to run target executables in
            if (++ixArg >= argc)
                Usage(argv[0], L"Missing arg for -s");
            if (0 == wcscmp(L"first", argv[ixArg]))
                whichSessions = WhichSessions_t::firstActive;
            else if (0 == wcscmp(L"active", argv[ixArg]))
                whichSessions = WhichSessions_t::allActive;
            else if (0 == wcscmp(L"all", argv[ixArg]))
                whichSessions = WhichSessions_t::allLoggedOn;
            else if (1 == swscanf_s(argv[ixArg], L"%lu", &nSessionId) && 0 != nSessionId)
                whichSessions = WhichSessions_t::oneSessionId;
            else
                Usage(argv[0], L"Invalid arg for -s", argv[ixArg]);
        }
        else if (0 == wcscmp(L"-32", argv[ixArg]))
        {
            // Keep WOW64 file system redirection; don't disable it.
            bWow64FileSystemRedir = true;
        }
        else if (0 == wcscmp(L"-q", argv[ixArg]))
        {
            // Quiet output
            bQuiet = true;
        }
        else if (0 == wcscmp(L"-debug", argv[ixArg]))
        {
            // Hidden debug option
            bDebug = true;
        }
        else if (0 == wcscmp(L"-debugF", argv[ixArg]))
        {
            // Hidden debug (file) option
            bDebug = bDebugF = true;
        }
        else
        {
            Usage(argv[0], L"Unrecognized command-line parameter", argv[ixArg]);
        }
        ++ixArg;
    }

    // A bit more command-line option validation
    if (bMergeStd && !bRedirStd)
    {
        Usage(argv[0], L"-merge is not valid without -redirStd");
    }

    // Redirection of standard out/err handles isn't effective if no wait time specified
    //TODO: should this be an error? Should it imply "-wait inf?"
    if (bRedirStd && 0 == dwWait)
    {
        bRedirStd = false;
        sRedirStdDirectory.clear();
        std::wcerr
            << L"Redirection of target process stdout/stderr is useful only when a wait time is specified with -wait or -term." << std::endl
            << L"Turning off stdout/stderr redirection." << std::endl;
    }

    // If sRedirStdDirectory is specified, ensure that it exists and is a directory, unless it's "-", in which case clear it.
    if (sRedirStdDirectory == L"-")
        sRedirStdDirectory.clear();
    if (sRedirStdDirectory.size() > 0)
    {
        // Remove trailing path separator if it has one. (PowerShell autocomplete likes to append them, helpfully...)
        while (EndsWith(sRedirStdDirectory, L'\\') || EndsWith(sRedirStdDirectory, L'/'))
            sRedirStdDirectory = sRedirStdDirectory.substr(0, sRedirStdDirectory.length() - 1);

        DWORD dwAttributes = GetFileAttributesW(sRedirStdDirectory.c_str());
        if (
            INVALID_FILE_ATTRIBUTES == dwAttributes ||
            0 == (FILE_ATTRIBUTE_DIRECTORY & dwAttributes)
            )
        {
            Usage(L"-redirStd argument is not a directory", argv[0]);
        }
    }

    // Implement hidden debug options
    if (bDebug)
    {
        // Write debug output to debug stream
        dbgOut.WriteToDebugStream(true);
        dbgOut.PrependTimestamp(true);
        if (bDebugF)
        {
            // and write debug output to a log file in the same directory with the executable.
            // Incorporate timestamp into filename
            std::wstringstream strDbgLogFname;
            wchar_t szPath[MAX_PATH * 2] = { 0 };
            GetModuleFileNameW(NULL, szPath, sizeof(szPath) / sizeof(szPath[0]));
            strDbgLogFname << szPath << L"." << TimestampUTCforFilepath(true) << L".log";
            sDbgLogFname = strDbgLogFname.str();
            dbgOut.WriteToFile(sDbgLogFname.c_str());
        }
    }

    // Prevent arithmetic overflow converting seconds to milliseconds.
    // If wait greater than about 49 days, make it infinite.
    if (dwWait >= 4294967)
        dwWait = INFINITE;
    // Convert seconds to milliseconds
    if (INFINITE != dwWait)
        dwWait *= 1000;

    //
    // Build the command line spec
    //
    if (!bPowerShell)
    {
        // Not using a PowerShell option - run the provided command line as is
        sActualCommandLine = sOriginalCommandLine;
    }
    else
    {
        // Run command line as a PowerShell command within a prepared PowerShell.exe environment.
        // 
        //TODO: Consider adding to the command to pick up the most likely "correct" exit code ($LASTEXITCODE); alternative is to rely on the command to have its own exit statement

        // Set up PowerShell command line to execute
        const std::wstring sPowerShellCmd = szPowerShellCmd;
        const std::wstring sEncodedCommand = L" -EncodedCommand ";
        const std::wstring sCommand = L" -Command ";

        if (bEncode)
        {
            // Base64-encode the input command line and run it with -EncodedCommand
            std::wstring sBase64Command;
            if (Base64Encode(sOriginalCommandLine, sBase64Command))
            {
                sActualCommandLine = sPowerShellCmd + sEncodedCommand + sBase64Command;
            }
            else
            {
                Usage(argv[0], L"Can't base64-encode command line");
            }
        }
        else if (bIsBase64)
        {
            // Input command line is base64-encoded; run it with -EncodedCommand
            sActualCommandLine = sPowerShellCmd + sEncodedCommand + sOriginalCommandLine;
        }
        else
        {
            // No encoding - just run the input command line with -Command
            sActualCommandLine = sPowerShellCmd + sCommand + sOriginalCommandLine;

            //TODO: Consider adding ampersand and curly braces after -Command, as described in the "powershell.exe /?" help text:
            /*
                -Command
                    Executes the specified commands (and any parameters) as though they were
                    typed at the Windows PowerShell command prompt, and then exits, unless
                    NoExit is specified. The value of Command can be "-", a string. or a
                    script block.

                    If the value of Command is "-", the command text is read from standard
                    input.

                    If the value of Command is a script block, the script block must be enclosed
                    in braces ({}). You can specify a script block only when running PowerShell.exe
                    in Windows PowerShell. The results of the script block are returned to the
                    parent shell as deserialized XML objects, not live objects.

                    If the value of Command is a string, Command must be the last parameter
                    in the command , because any characters typed after the command are
                    interpreted as the command arguments.

                    To write a string that runs a Windows PowerShell command, use the format:
                        "& {<command>}"
                    where the quotation marks indicate a string and the invoke operator (&)
                    causes the command to be executed.
            */
        }
    }

    if (!bQuiet)
    {
        std::wcout << L"Command line : " << sActualCommandLine << std::endl;
        if (bPowerShell)
            std::wcout << L"  Originally : " << sOriginalCommandLine << std::endl;
        std::wcout << L"Sessions     ? " << WhichSessionsToWSZ(whichSessions, nSessionId) << std::endl;
        std::wcout << L"PowerShell   ? " << (bPowerShell ? L"Yes" : L"No") << std::endl;
        std::wcout << L"Redir Std    ? ";
        if (!bRedirStd)
        {
            std::wcout << L"Not redirected" << std::endl;
        }
        else
        {
            if (sRedirStdDirectory.length() > 0)
                std::wcout << sRedirStdDirectory << std::endl;
            else
                std::wcout << L"This process" << std::endl;
            if (bMergeStd)
                std::wcout << L"               Merging targets' stderr into stdout" << std::endl;
            else
                std::wcout << L"               Keeping targets' stderr and stdout separate" << std::endl;
        }
        std::wcout << L"Try elevated ? " << (bTryElevated ? L"Yes" : L"No") << std::endl;
        std::wcout << L"WOW64 redir  ? " << (bWow64FileSystemRedir ? L"Enabled" : L"Disabled") << std::endl;
        std::wcout << L"Hidden       ? " << (bHidden ? L"Yes" : L"No") << std::endl;
        std::wcout << L"Minimized    ? " << (bMinimized ? L"Yes" : L"No") << std::endl;
        if (0 == dwWait)
        {
            std::wcout << L"Wait         ? " << L"No" << std::endl;
        }
        else
        {
            if (INFINITE == dwWait)
                std::wcout << L"Wait         ? " << L"Yes: infinite." << std::endl;
            else
                std::wcout << L"Wait         ? " << L"Yes: " << dwWait << L" milliseconds." << std::endl;
            std::wcout << L"Terminate    ? " << (bTerminate ? L"Yes" : L"No") << std::endl;
        }
        if (bDebug)
        {
            if (bDebugF)
                std::wcout << L"Debug output : debug stream and " << sDbgLogFname << std::endl;
            else
                std::wcout << L"Debug output : debug stream" << std::endl;
        }
        std::wcout << std::endl;
    }

    // Must be running as System
    // (Could perform this check at the start, but this way the operator can test out command line parsing without having to run as System.)
    WhoAmI whoAmI;
    if (!whoAmI.IsSystem())
    {
        std::wcerr << L"ERROR: This program must be executed as SYSTEM" << std::endl;
        exit(-2);
    }

    // Start by getting info on all WTS sessions
    PWTS_SESSION_INFOW pSessionInfo = NULL;
    DWORD dwSessionCount = 0;
    BOOL ret = WTSEnumerateSessionsW(WTS_CURRENT_SERVER_HANDLE, 0, 1, &pSessionInfo, &dwSessionCount);
    if (!ret)
    {
        dwLastErr = GetLastError();
        std::wcerr << L"Cannot enumerate WTS sessions: " << SysErrorMessageWithCode(dwLastErr) << std::endl;
        exit(-3);
    }

    // Instantiate an object to handle processes that get launched
    ProcessManager_t processManager;

    bool bDoneWithSessions = false;
    for (DWORD ixSession = 0; ixSession < dwSessionCount && !bDoneWithSessions; ++ixSession)
    {
        // Explicitly skip session 0
        if (0 == pSessionInfo[ixSession].SessionId)
        {
            continue;
        }

        // Add a ptrSessionProcessInfo_t to the collection, even if we end up not launching a process in this session
        ptrSessionProcessInfo_t pSPI = processManager.New();
        pSPI->session.dwSessionId = pSessionInfo[ixSession].SessionId;
        pSPI->session.wtsState = pSessionInfo[ixSession].State;

        if (!bQuiet)
            std::wcout << L"Session ID " << pSPI->session.dwSessionId << L", state: " << WtsConnectStateToWSZ(pSPI->session.wtsState) << L"; WinSta name: " << pSessionInfo[ixSession].pWinStationName << std::endl;

        // Get more info about this session, including user domain\name, logon time, and whether the session is locked.
        LPWSTR pInfo = nullptr;
        DWORD dwBytesReturned = 0;
        ret = WTSQuerySessionInformationW(WTS_CURRENT_SERVER_HANDLE, pSPI->session.dwSessionId, WTSSessionInfoEx, &pInfo, &dwBytesReturned);
        if (ret)
        {
            const WTSINFOEXW* pWtsInfo = (const WTSINFOEXW*)pInfo;
            const WTSINFOEX_LEVEL1_W& wtsInfo = pWtsInfo->Data.WTSInfoExLevel1;
            pSPI->session.sDomain = wtsInfo.DomainName;
            pSPI->session.sUser = wtsInfo.UserName;
            pSPI->session.wtsFlags = wtsInfo.SessionFlags;
            pSPI->session.logonTime = wtsInfo.LogonTime;
            WTSFreeMemory(pInfo);

            if (!bQuiet && pSPI->session.sUser.length() > 0)
            {
                // Report more information about the session.
                // Win7/WS2008R2 reports lock state incorrectly, so don't include that text if running on W7/WS2008R2
                if (!IsWin7orWS2008R2())
                    std::wcout << L"User " << pSPI->session.sDomain << L"\\" << pSPI->session.sUser << L", logon time " << LargeIntegerToDateTimeString(pSPI->session.logonTime) << L", session " << WtsFlagsToWSZ(pSPI->session.wtsFlags) << std::endl;
                else
                    std::wcout << L"User " << pSPI->session.sDomain << L"\\" << pSPI->session.sUser << L", logon time " << LargeIntegerToDateTimeString(pSPI->session.logonTime) << std::endl;
            }
        }
        else
        {
            // For cases where WTSSessionInfoEx isn't supported
            ret = WTSQuerySessionInformationW(WTS_CURRENT_SERVER_HANDLE, pSPI->session.dwSessionId, WTSSessionInfo, &pInfo, &dwBytesReturned);
            if (ret)
            {
                const WTSINFOW* pWtsInfo = (const WTSINFOW*)pInfo;
                pSPI->session.sDomain = pWtsInfo->Domain;
                pSPI->session.sUser = pWtsInfo->UserName;
                pSPI->session.logonTime = pWtsInfo->LogonTime;
                WTSFreeMemory(pInfo);

                if (!bQuiet && pSPI->session.sUser.length() > 0)
                    std::wcout << L"User " << pSPI->session.sDomain << L"\\" << pSPI->session.sUser << L", logon time " << LargeIntegerToDateTimeString(pSPI->session.logonTime) << std::endl;
            }
            else
            {
                dwLastErr = GetLastError();
                std::wcerr << L"Could not retrieve session information: " << SysErrorMessageWithCode(dwLastErr) << std::endl;
            }
        }

        // Start process in this session, depending on the "whichSessions" setting;
        // Set flag to exit loop if only one session to be targeted.
        bool bStartProcessInThisSession = false, bExitLoopAfterThisOne = false;
        switch (whichSessions)
        {
        case WhichSessions_t::firstActive:
            // Start process if this session is active, then exit the loop.
            if (WTSActive == pSPI->session.wtsState)
            {
                bStartProcessInThisSession = bExitLoopAfterThisOne = true;
            }
            break;
        case WhichSessions_t::allActive:
            // Start process if this session is active
            bStartProcessInThisSession = (WTSActive == pSPI->session.wtsState);
            break;
        case WhichSessions_t::allLoggedOn:
            // Start process if this session is active or disconnected
            bStartProcessInThisSession = (WTSActive == pSPI->session.wtsState || WTSDisconnected == pSPI->session.wtsState);
            break;
        case WhichSessions_t::oneSessionId:
            // Start process if session ID matches the one specified AND the session is active or disconnected.
            if (nSessionId == pSPI->session.dwSessionId)
            {
                // Session ID matches, so no need to continue enumeration
                bExitLoopAfterThisOne = true;
                if (WTSActive == pSPI->session.wtsState || WTSDisconnected == pSPI->session.wtsState)
                {
                    bStartProcessInThisSession = true;
                }
                else
                {
                    std::wcerr << L"Session ID " << nSessionId << L" exists but is not active or disconnected." << std::endl;
                }
            }
            break;
        }
        if (bStartProcessInThisSession)
        {
            nTargetedSessions++;

            // Get the user token associated with the session
            HANDLE hToken = NULL;
            ret = WTSQueryUserToken(pSPI->session.dwSessionId, &hToken);
            if (!ret)
            {
                dwLastErr = GetLastError();
                std::wcerr << L"Cannot query user token: " << SysErrorMessageWithCode(dwLastErr) << std::endl;
            }
            else
            {
                // If the try-elevated option is set and there's a higher-IL linked token, get it.
                if (bTryElevated && Token::GetHighestToken(hToken))
                {
                    pSPI->process.bElevated = true;
                }

                // Turn off WOW64 file system redirection by default (no-op if this is a 64-bit process or a 32-bit OS)
                Wow64FsRedirection fsRedir;
                if (!bWow64FileSystemRedir)
                    fsRedir.Disable();

                // Create the appropriate environment block for this user
                LPVOID pEnv = nullptr;
                ret = CreateEnvironmentBlock(&pEnv, hToken, FALSE);
                if (ret)
                {
                    PROCESS_INFORMATION pi = { 0 };
                    STARTUPINFOW si = { 0 };
                    HANDLE hPipeStdoutWr = NULL, hPipeStderrWr = NULL, hPipeStdinWr = NULL, hPipeStdinRd = NULL;

                    si.cb = sizeof(si);
                    // Implement hidden/minimized options
                    if (bHidden || bMinimized)
                    {
                        si.dwFlags = STARTF_USESHOWWINDOW;
                        si.wShowWindow = bHidden ? SW_HIDE : SW_SHOWMINNOACTIVE;
                    }

                    if (bRedirStd)
                    {
                        // To redirect the child process' stdout and stderr, create anonymous pipes for stdin/stdout/stderr,
                        // with handles for the child process marked inheritable, and provide those handles to the
                        // child process through the STARTUPINFOW structure.
                        // Hold onto handles for the "read" ends of the stdout and stderr pipes.
                        // We're not doing anything with stdin but we need to provide all three.
                        // If merging stderr with stdout, create just one pipe for both and provide its write handle for
                        // both stdout and stderr.

                        // Security attributes: inheritable
                        SECURITY_ATTRIBUTES sa = { 0 };
                        sa.nLength = sizeof(sa);
                        sa.bInheritHandle = TRUE;
                        sa.lpSecurityDescriptor = NULL;

                        // Default pipe size; we will be monitoring those pipes continually.
                        // Create pipes for stdin and stdout; create a separate pipe for stderr unless we're merging
                        // stderr into stdout.
                        BOOL bPipesCreated =
                            (FALSE != CreatePipe(&hPipeStdinRd, &hPipeStdinWr, &sa, 0)) &&
                            (FALSE != CreatePipe(&pSPI->process.hPipeStdoutRd, &hPipeStdoutWr, &sa, 0));
                        if (bPipesCreated && !bMergeStd && !CreatePipe(&pSPI->process.hPipeStderrRd, &hPipeStderrWr, &sa, 0))
                            bPipesCreated = FALSE;
                        if (!bPipesCreated)
                        {
                            dwLastErr = GetLastError();
                            std::wcerr << L"Error building pipe; " << SysErrorMessageWithCode(dwLastErr) << std::endl;
                            //TODO: How should CreatePipe errors be handled? Terminate the process?
                            exit(-3);
                        }

                        // We created the pipes with inheritable handles, as we want the child process to inherit
                        // the stdin "read" and the stdout/stderr "write" handles. We don't want the child process
                        // to inherit the remaining handles, so clear the "inheritable" flag on the handles for 
                        // the other ends of those pipes.
                        // (hPipeStderrRd will be NULL if we're using the stdout pipe for both stdout and stderr.)
                        SetHandleInformation(pSPI->process.hPipeStdoutRd, HANDLE_FLAG_INHERIT, 0);
                        if (NULL != pSPI->process.hPipeStderrRd)
                            SetHandleInformation(pSPI->process.hPipeStderrRd, HANDLE_FLAG_INHERIT, 0);
                        SetHandleInformation(hPipeStdinWr, HANDLE_FLAG_INHERIT, 0);

                        // Set the standard handles for the new process
                        si.dwFlags |= STARTF_USESTDHANDLES;
                        si.hStdInput = hPipeStdinRd;
                        si.hStdOutput = hPipeStdoutWr;
                        // Choice whether to redirect child process' stderr to the same pipe that stdout is going to
                        si.hStdError = bMergeStd ? hPipeStdoutWr : hPipeStderrWr;
                    }

                    // CreateProcessAsUserW third parameter is supposed to be a non-const buffer; casting a const pointer to non-const is ungood.
                    // Create a new buffer and fill it with null characters, then copy string into it.
                    size_t cmdLineBufSize = sActualCommandLine.length() + 1;
                    wchar_t* szActualCommandLine = new wchar_t[cmdLineBufSize] { 0 };
                    sActualCommandLine._Copy_s(szActualCommandLine, cmdLineBufSize, sActualCommandLine.length());
                    // Clear the last error prior to invoking the API, in case there's a failure code path that doesn't set the thread's last error value.
                    SetLastError(0);
                    // Start the target process; the token specifies the WTS session in which the process will execute.
                    // Start it with the primary thread suspended until after we set up any required redirection.
                    ret = CreateProcessAsUserW(
                        hToken, 
                        nullptr,
                        szActualCommandLine,
                        nullptr, 
                        nullptr, 
                        TRUE, 
                        CREATE_BREAKAWAY_FROM_JOB | CREATE_NEW_CONSOLE | CREATE_UNICODE_ENVIRONMENT | CREATE_SUSPENDED,
                        pEnv, 
                        TargetCurrentDirectory().c_str(),
                        &si, 
                        &pi);
                    dwLastErr = GetLastError();
                    // And delete that buffer now
                    delete[] szActualCommandLine;
                    if (ret)
                    {
                        // Get info about the new process.
                        pSPI->process.hProcess = pi.hProcess;
                        pSPI->process.dwPID = pi.dwProcessId;
                        // Set up redirection and the monitoring of stdout/stderr
                        SetUpRedirection(pSPI, bRedirStd, bMergeStd, sRedirStdDirectory);
                        ResumeThread(pi.hThread);
                        CloseHandle(pi.hThread);

                        if (!bQuiet)
                            std::wcout << L"PID " << pSPI->process.dwPID << L" started in session " << pSPI->session.dwSessionId << L" running " << (pSPI->process.bElevated ? L"elevated" : L"non-elevated") << L" as " << pSPI->session.sDomain << L"\\" << pSPI->session.sUser << std::endl;
                    }
                    else
                    {
                        std::wcerr << L"CreateProcessAsUserW failed: " << SysErrorMessageWithCode(dwLastErr) << std::endl;
                    }
                    if (bRedirStd)
                    {
                        // Close the pipe handles we no longer need - the ones inherited by the child process, 
                        // and the stdin "write" handle as we're not writing to its stdin.
                        CloseHandle(hPipeStdoutWr);
                        CloseHandle(hPipeStderrWr);
                        CloseHandle(hPipeStdinWr);
                        CloseHandle(hPipeStdinRd);
                    }
                    // Deallocate allocated memory
                    DestroyEnvironmentBlock(pEnv);

                    // Restore previous WOW64 file system redirection state
                    if (!bWow64FileSystemRedir)
                        fsRedir.Revert();
                }
                else
                {
                    dwLastErr = GetLastError();
                    std::wcerr << L"Cannot create environment block for user: " << SysErrorMessageWithCode(dwLastErr) << std::endl;
                }

                CloseHandle(hToken);
            }

            // If "first active" only, then exit the loop
            if (bExitLoopAfterThisOne)
            {
                if (!bQuiet)
                {
                    switch (whichSessions)
                    {
                    case WhichSessions_t::firstActive:
                        std::wcout << L"First active session processed; not looking at any more sessions." << std::endl;
                        break;
                    case WhichSessions_t::oneSessionId:
                        std::wcout << L"Session " << nSessionId << L" processed; not looking at any more sessions." << std::endl;
                        break;
                    case WhichSessions_t::allActive:
                    case WhichSessions_t::allLoggedOn:
                        // These are never the case when exiting loop "after this one" 
                        // but the compiler wants all enums handled in the switch
                        break;
                    }
                }
                bDoneWithSessions = true;
            }
        }

        if (!bQuiet)
            std::wcout << std::endl;
    }

    WTSFreeMemory(pSessionInfo);
    pSessionInfo = nullptr;

    // If waiting for processes, start waiting
    if (0 != dwWait)
    {
        ULARGE_INTEGER ulStartTime;
        GetSystemTimeAsULargeinteger(ulStartTime);
        DWORD dwNextWait = dwWait;

        bool bKeepMonitoring = true;
        while (bKeepMonitoring)
        {
            // Wait for any of the launched processes to exit; report exit code.
            // Function returns how many processes were running, how many are still running, and which processes (if any) exited.
            DWORD nRunningProcesses = 0, nNowRunning = 0;
            vecSessionProcessInfo_t v_ExitedProcesses;
            bool bAnyExited = processManager.WaitForAProcessToExit(dwNextWait, nRunningProcesses, nNowRunning, v_ExitedProcesses);
            if (bAnyExited)
            {
                // Report exit code for those that exited.
                for (auto iter = v_ExitedProcesses.begin(); iter != v_ExitedProcesses.end(); ++iter)
                {
                    const ptrSessionProcessInfo_t& pSPI = *iter;
                    std::wcout << L"Process " << pSPI->process.dwPID << L" running as " << pSPI->session.sUser << L" in session " << pSPI->session.dwSessionId << L" exited; exit code " << pSPI->process.dwExitCode << std::endl;
                }
            }
            else
            {
                //std::wcout << L"No processes exited; " << nRunningProcesses << L" left." << std::endl;
            }
            // Stop monitoring if none of the launched processes are still running.
            bKeepMonitoring = (nNowRunning > 0);
            if (bKeepMonitoring)
            {
                // Calculate remaining wait time. If "INFINITE" it's still infinite.
                if (INFINITE != dwWait)
                {
                    // How long since we started monitoring
                    DWORD dwMsSince = MillisecondsSince(ulStartTime);
                    if (dwWait > dwMsSince)
                    {
                        // If there's time left, that's the new wait time.
                        dwNextWait = dwWait - dwMsSince;
                    }
                    else
                    {
                        // No time remaining; stop the remaining monitors, and optionally terminate the processes.
                        if (bTerminate)
                            std::wcout << L"Timeout expired; terminating " << nNowRunning << L" remaining process(es)" << std::endl;
                        else
                            std::wcout << L"Timeout expired; " << nNowRunning << L" process(es) still running" << std::endl;
                        bKeepMonitoring = false;
                        processManager.StopAllRedirectionMonitors(bTerminate);
                    }
                }
            }
        }

        // Wait for the redirection monitors to exit before allowing processManager and other objects to go
        // out of scope, deallocating global objects, etc.
        processManager.WaitForRedirectionMonitors();
    }

    if (0 == nTargetedSessions)
    {
        std::wcout << L"No sessions matching specified criteria." << std::endl;
    }

    return 0;
}

// ----------------------------------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------------------------------
