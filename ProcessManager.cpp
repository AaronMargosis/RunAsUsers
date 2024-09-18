// Code to manage processes started in users' sessions

#include <iostream>
#include "ProcessManager.h"
#include "SysErrorMessage.h"

#include "DbgOut.h"

// ------------------------------------------------------------------------------------------

/// <summary>
/// Release all resources
/// </summary>
void ProcessInfo_t::Uninit()
{
    CloseHandle(hProcess);
    CloseHandle(hPipeStdoutRd);
    CloseHandle(hPipeStderrRd);
    // Don't close the redir target if it's this process' stdout handle
    if (GetStdHandle(STD_OUTPUT_HANDLE) != hStdoutRedirTarget)
        CloseHandle(hStdoutRedirTarget);
    // Don't close the redir target if it's this process' stderr handle
    if (GetStdHandle(STD_ERROR_HANDLE) != hStderrRedirTarget)
        CloseHandle(hStderrRedirTarget);
    CloseHandle(hThread_StdoutMonitor);
    CloseHandle(hThread_StderrMonitor);

    hProcess = 
        hPipeStdoutRd = 
        hPipeStderrRd = 
        hStdoutRedirTarget = 
        hStderrRedirTarget = 
        hThread_StdoutMonitor = 
        hThread_StderrMonitor = NULL;
}

// ------------------------------------------------------------------------------------------

void ProcessManager_t::Clear()
{
    // Empty the collection. When objects' reference counts hit zero, they will be cleaned up.
    m_processes.clear();
}

/// <summary>
/// Waits up to dwTimeout milliseconds for one or more still-running processes in the collection to exit.
/// Changes the bExited status of those processes and returns information about them in the vExitedProcesses
/// parameter.
/// </summary>
/// <param name="dwTimeout">Input: number of milliseconds to wait, or INFINITE to wait indefinitely.</param>
/// <param name="nRunningProcesses">Output: number of processes that were running at the beginning of the wait.</param>
/// <param name="nNowRunning">Output: number of processes that were still running at the end of the wait.</param>
/// <param name="vExitedProcesses">Output: collection of processes that exited during this wait.</param>
/// <returns>true if one or more processes exited during this wait; false otherwise</returns>
bool ProcessManager_t::WaitForAProcessToExit(DWORD dwTimeout, DWORD& nRunningProcesses, DWORD& nNowRunning, vecSessionProcessInfo_t& vExitedProcesses)
{
    bool retval = false;
    vExitedProcesses.clear();
    nRunningProcesses = nNowRunning = 0;

    dbgOut.locked() << L"WaitForAProcessToExit, timeout " << dwTimeout << std::endl;

    if (m_processes.size() == 0)
        return false;

    // Use std::vector to create a contiguous array of still-running process handles to pass to WaitForMultipleObjects
    std::vector<HANDLE> vHProcesses;;
    for (auto iter = ConstIter(); !IterAtEnd(iter); iter++)
    {
        // Add handles of running processes to the array
        const ptrSessionProcessInfo_t& pSPI = *iter;
        if (NULL != pSPI->process.hProcess && !pSPI->process.bExited)
        {
            vHProcesses.push_back(pSPI->process.hProcess);
            nRunningProcesses++;
        }
    }

    dbgOut.locked() << L"nRunningProcesses = " << nRunningProcesses << std::endl;
    
    if (nRunningProcesses > 0)
    {
        // Wait for any of the processes to exit
        DWORD wfmoRet = WaitForMultipleObjects(nRunningProcesses, vHProcesses.data(), FALSE, dwTimeout);
        if (WAIT_TIMEOUT == wfmoRet)
        {
            // None exited during the timeout period. Update the nNowRunning return parameter.
            dbgOut.locked() << L"No processes exited during the timeout period" << std::endl;
            nNowRunning = nRunningProcesses;
        }
        else
        {
            // One or more exited; identify them
            for (auto iter = Iter(); !IterAtEnd(iter); iter++)
            {
                // Look only at the processes that had been running this time through
                ptrSessionProcessInfo_t& pSPI = *iter;
                if (NULL != pSPI->process.hProcess && !pSPI->process.bExited)
                {
                    // Use WaitForSingleObject with a timeout of 0 to determine whether a specific process has exited.
                    // Can't use GetExitCodeProcess to determine whether a process has exited:
                    // it doesn't distinguish between a process that has exited and one that exited with 259 (STILL_ACTIVE)
                    DWORD wfsoRet = WaitForSingleObject(pSPI->process.hProcess, 0);
                    switch (wfsoRet)
                    {
                    case WAIT_OBJECT_0:
                        // The process has exited; get its exit code
                        GetExitCodeProcess(pSPI->process.hProcess, &pSPI->process.dwExitCode);
                        // add it to the vExitedProcesses output collection
                        vExitedProcesses.push_back(*iter);
                        // Flag it as having exited so we don't look at it again
                        pSPI->process.bExited = true;
                        break;

                    case WAIT_TIMEOUT:
                        // process is still running
                        ++nNowRunning;
                        break;

                    default:
                        //TODO: how to handle this error (unexpected return value from WaitForSingleObject)
                        {
                            DWORD dwLastErr = GetLastError();
                            std::wcerr << L"WaitForSingleObject for PID " << pSPI->process.dwPID << L" failed: " << SysErrorMessageWithCode(dwLastErr) << std::endl;
                        }
                        break;
                    }
                }
            }

            retval = true;
        }
    }

    return retval;
}

/// <summary>
/// For use when monitoring timeout expires: stop all threads monitoring pipes for redirected output, 
/// and optionally terminate still-running processes.
/// </summary>
/// <param name="bTerminateProcesses">Input: whether to terminate still-running processes</param>
void ProcessManager_t::StopAllRedirectionMonitors(bool bTerminateProcesses)
{
    // For each process in the collection:
    // cancel any blocked I/O operations; this will cause ReadFile to fail with an
    // error that we recognize.
    // Terminate the process if bTerminateProcesses is true.
    for (auto iter = Iter(); !IterAtEnd(iter); iter++)
    {
        // Only either CancelIoEx or CancelSynchronousIo is really needed. Doing both anyway.
        const ptrSessionProcessInfo_t& pSPI = *iter;
        if (NULL != pSPI->process.hPipeStdoutRd)
            CancelIoEx(pSPI->process.hPipeStdoutRd, NULL);
        if (NULL != pSPI->process.hPipeStderrRd)
            CancelIoEx(pSPI->process.hPipeStderrRd, NULL);
        if (NULL != pSPI->process.hThread_StdoutMonitor)
            CancelSynchronousIo(pSPI->process.hThread_StdoutMonitor);
        if (NULL != pSPI->process.hThread_StderrMonitor)
            CancelSynchronousIo(pSPI->process.hThread_StderrMonitor);

        if (bTerminateProcesses)
            TerminateProcess(pSPI->process.hProcess, ERROR_TIMEOUT);
    }
}

/// <summary>
/// Wait for all threads monitoring redirected output to exit
/// </summary>
void ProcessManager_t::WaitForRedirectionMonitors()
{
    // Use std::vector to create a contiguous array of thread handles to pass to WaitForMultipleObjects
    std::vector<HANDLE> vHThreads;
    // Count threads in a DWORD to avoid type mismatch between the DWORD that WaitForMultipleObjects wants
    // and std::vector size() which returns a size_t.
    DWORD nThreads = 0;
    for (auto iter = ConstIter(); !IterAtEnd(iter); iter++)
    {
        // Add each valid thread handle to the array
        const ptrSessionProcessInfo_t& pSPI = *iter;
        if (NULL != pSPI->process.hThread_StdoutMonitor)
        {
            vHThreads.push_back(pSPI->process.hThread_StdoutMonitor);
            nThreads++;
        }
        if (NULL != pSPI->process.hThread_StderrMonitor)
        {
            vHThreads.push_back(pSPI->process.hThread_StderrMonitor);
            nThreads++;
        }
    }
    // Wait for all handles to go into the signalled state (indicating the threads have all exited)
    WaitForMultipleObjects(nThreads, vHThreads.data(), TRUE, INFINITE);
}

