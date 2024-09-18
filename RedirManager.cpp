#include <Windows.h>
#include <iostream>
#include <sstream>
#include "StringUtils.h"
#include "UtilityFunctions.h"
#include "SysErrorMessage.h"
#include "RedirManager.h"
#include "DbgOut.h"


/// <summary>
/// Function to copy data from a named or anonymous pipe to a file object.
/// </summary>
/// <param name="hPipe">Handle to a named or anonymous pipe to read from</param>
/// <param name="hDestination">Handle to a file object to write to</param>
/// <param name="pSPI">Information about the session/process providing the data (for debugging/diagnostic purposes)</param>
/// <returns>(Always returns true in current implementation)</returns>
static bool ReadPipeToFile(HANDLE hPipe, HANDLE hDestination, ptrSessionProcessInfo_t pSPI)
{
    // Allocate buffer: read up to 1MB at a time
    const size_t cbBufSize = 1024 * 1024;
    uint8_t* buffer = new uint8_t[cbBufSize];

    DWORD dwPID = pSPI->process.dwPID;

    bool bContinue = true;
    while (bContinue)
    {
        DWORD dwRead = 0;
        SecureZeroMemory(buffer, cbBufSize);
        // Read from the pipe until it's empty
        BOOL rfRet = ReadFile(hPipe, buffer, cbBufSize - 8, &dwRead, NULL);
        DWORD dwLastErr = GetLastError();
        if (rfRet)
        {
            // ReadFile succeeded for PID dwPID; read dwRead bytes
            dbgOut.locked() << L"ReadPipeToFile for PID " << dwPID << L"; ReadFile read " << dwRead << L" bytes" << std::endl;
        }
        else if (ERROR_BROKEN_PIPE == dwLastErr)
        {
            // ReadFile failed with ERROR_BROKEN_PIPE: should be good now
            dbgOut.locked() << L"ReadPipeToFile for PID " << dwPID << L" ERROR_BROKEN_PIPE - should be good now" << std::endl;
        }
        else if (ERROR_OPERATION_ABORTED == dwLastErr)
        {
            // ReadFile failed with ERROR_OPERATION_ABORTED: time must be up
            dbgOut.locked() << L"ReadPipeToFile for PID " << dwPID << L" ERROR_OPERATION_ABORTED - time must be up" << std::endl;
        }
        else
            std::wcerr << L"ReadFile error with PID " << dwPID << L": " << SysErrorMessageWithCode(dwLastErr) << std::endl;
        bContinue = (rfRet && dwRead > 0);
        if (bContinue)
        {
            DWORD dwWritten = 0;
            BOOL ret = WriteFile(hDestination, buffer, dwRead, &dwWritten, NULL);
            if (ret)
            {
                if (dwRead != dwWritten)
                    std::wcerr << L"WriteFile anomaly: read " << dwRead << L" bytes but wrote " << dwWritten << std::endl;
            }
            else
            {
                dwLastErr = GetLastError();
                std::wcerr << L"WriteFile error: " << SysErrorMessageWithCode(dwLastErr) << std::endl;
            }
        }
    }
    delete[] buffer;
    return true;
}

/// <summary>
/// Thread function to monitor pipe for the process' stderr output.
/// </summary>
/// <param name="lpvThreadParameter">Input: pointer to ptrSessionProcessInfo_t returned by CreateCrossThreadpSPI, referencing the process to monitor</param>
/// <returns></returns>
static DWORD WINAPI StderrMonitor(LPVOID lpvThreadParameter)
{
    // Use ConsumeCrossThreadpSPI to properly handle object passed via CreateThread.
    ptrSessionProcessInfo_t pSPI;
    ConsumeCrossThreadpSPI(pSPI, lpvThreadParameter);

    dbgOut.locked() << L"StderrMonitor start for PID " << pSPI->process.dwPID << std::endl;

    bool ret = ReadPipeToFile(pSPI->process.hPipeStderrRd, pSPI->process.hStderrRedirTarget, pSPI);

    dbgOut.locked() << L"StderrMonitor exit for PID " << pSPI->process.dwPID << L"; ReadPipeToFile returned " << ret << std::endl;

    return 0;
}

/// <summary>
/// Thread function to monitor pipe for the process' stdout output. If its stderr needs separate monitoring, starts a thread for that.
/// </summary>
/// <param name="lpvThreadParameter">Input: pointer to ptrSessionProcessInfo_t returned by CreateCrossThreadpSPI, referencing the process to monitor</param>
/// <returns></returns>
static DWORD WINAPI StdoutMonitor(LPVOID lpvThreadParameter)
{
    // Use ConsumeCrossThreadpSPI to properly handle object passed via CreateThread.
    ptrSessionProcessInfo_t pSPI;
    ConsumeCrossThreadpSPI(pSPI, lpvThreadParameter);

    // If its stderr needs separate monitoring, start a thread for that.
    if (NULL != pSPI->process.hPipeStderrRd && NULL != pSPI->process.hStderrRedirTarget)
    {
        // Use CreateCrossThreadpSPI to get an address of a ptrSessionProcessInfo_t that is safe to pass via CreateThread.
        pSPI->process.hThread_StderrMonitor = CreateThread(NULL, 0, StderrMonitor, CreateCrossThreadpSPI(pSPI), 0, NULL);
    }

    bool ret = ReadPipeToFile(pSPI->process.hPipeStdoutRd, pSPI->process.hStdoutRedirTarget, pSPI);

    dbgOut.locked() << L"StdoutMonitor exit for PID " << pSPI->process.dwPID << L"; ReadPipeToFile returned " << ret << std::endl;

    return 0;
}

/// <summary>
/// Sets up everything for redirecting a target process' stdout/stderr to a destination
/// </summary>
/// <param name="pSPI">ptrSessionProcessInfo_t for the process to monitor</param>
/// <param name="bRedirStd">Input: whether to redirect stdout/stderr</param>
/// <param name="bMergeStd">Input: whether to merge stderr into stdout</param>
/// <param name="sRedirStdDirectory">Input: directory in which to create output file(s) for redirection; if empty, then redirect to this process' stdout/stderr</param>
/// <returns></returns>
bool SetUpRedirection(ptrSessionProcessInfo_t& pSPI, bool bRedirStd, bool bMergeStd, const std::wstring& sRedirStdDirectory)
{
    // Nothing to do
	if (!bRedirStd)
		return true;

    // Redirecting stdout/stderr to file(s) in the named directory
	if (sRedirStdDirectory.length() > 0)
	{
        // Define file names for stdout and stderr (use only the former if stderr is merged into stdout).
        // Incorporate session number, process ID, and timestamp into the filename.
		//TODO: Put the timestamp at the beginning of the filename? Maybe put something about the command line in the file name? Get user name into the file name?
		std::wstringstream strFnameStdout, strFnameStderr;
		std::wstring sTimestampForFilename = TimestampUTCforFilepath(false);
		strFnameStdout << sRedirStdDirectory << L"\\S_" << pSPI->session.dwSessionId << L"_P_" << pSPI->process.dwPID << L"_stdout_" << sTimestampForFilename << L".txt";
		strFnameStderr << sRedirStdDirectory << L"\\S_" << pSPI->session.dwSessionId << L"_P_" << pSPI->process.dwPID << L"_stderr_" << sTimestampForFilename << L".txt";

        // Create the file where redirected stdout goes
		pSPI->process.hStdoutRedirTarget = CreateFileW(
			strFnameStdout.str().c_str(), 
			GENERIC_WRITE, 
			FILE_SHARE_READ, 
			NULL, 
			CREATE_ALWAYS, 
			FILE_ATTRIBUTE_NORMAL, 
			NULL);
		if (INVALID_HANDLE_VALUE == pSPI->process.hStdoutRedirTarget)
		{
			DWORD dwLastErr = GetLastError();
			std::wcerr << L"CreateFileW failed for " << strFnameStdout.str() << L": " << SysErrorMessageWithCode(dwLastErr) << std::endl;
		}

        // If stderr isn't redirected into stdout, create the file where redirected stderr goes
		if (!bMergeStd)
		{
			pSPI->process.hStderrRedirTarget = CreateFileW(
				strFnameStderr.str().c_str(),
				GENERIC_WRITE,
				FILE_SHARE_READ,
				NULL,
				CREATE_ALWAYS,
				FILE_ATTRIBUTE_NORMAL,
				NULL);
			if (INVALID_HANDLE_VALUE == pSPI->process.hStderrRedirTarget)
			{
				DWORD dwLastErr = GetLastError();
				std::wcerr << L"CreateFileW failed for " << strFnameStderr.str() << L": " << SysErrorMessageWithCode(dwLastErr) << std::endl;
			}
		}
	}
	else
	{
        // Not redirecting to file(s) - redirect to this process' stdout/stderr
		pSPI->process.hStdoutRedirTarget = GetStdHandle(STD_OUTPUT_HANDLE);
		if (!bMergeStd)
			pSPI->process.hStderrRedirTarget = GetStdHandle(STD_ERROR_HANDLE);
	}

    // Start the thread to monitor the stdout pipe; if necessary it will start another thread to monitor the stderr pipe.
    // Use CreateCrossThreadpSPI to get an address of a ptrSessionProcessInfo_t that is safe to pass via CreateThread.
    pSPI->process.hThread_StdoutMonitor = CreateThread(NULL, 0, StdoutMonitor, CreateCrossThreadpSPI(pSPI), 0, NULL);

	return true;
}
