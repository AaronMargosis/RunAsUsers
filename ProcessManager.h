// Code to manage processes started in users' sessions

#pragma once

#include <Windows.h>
#include <WtsApi32.h>
#include <string>
#include <vector>
#include <memory>


/// <summary>
/// Session-specific information
/// </summary>
struct SessionInfo_t
{
    DWORD dwSessionId = 0;
    WTS_CONNECTSTATE_CLASS wtsState = WTS_CONNECTSTATE_CLASS::WTSInit; // Compiler wants it initialized explicitly to something
    std::wstring sDomain;
    std::wstring sUser;
    // From WTSINFOEX_LEVEL1_W SessionFlags
    LONG wtsFlags = 0;
    LARGE_INTEGER logonTime = { 0 };

public:
    // Explicitly declare ctor/dtor
    SessionInfo_t() = default;
    ~SessionInfo_t() = default;

private:
    // Explicitly delete cctor/assignment
    SessionInfo_t(const SessionInfo_t&) = delete;
    SessionInfo_t& operator = (const SessionInfo_t&) = delete;
};

/// <summary>
/// Process-specific information
/// </summary>
struct ProcessInfo_t
{
    // This object owns all its HANDLE values and is responsible for closing them when no longer needed.

    // Handle to the process
    HANDLE hProcess = NULL;
    // process ID
    DWORD dwPID = 0;
    // set to true after the process has exited
    bool bExited = false;
    // Exit code valid only when bExited is true.
    DWORD dwExitCode = 0;
    // Whether the process is running elevated
    bool bElevated = false;

    // read handles for each process' redirected stdout and stderr
    // Both are NULL if not redirecting the process' stdout/stderr
    // If hPipeStdoutRd is non-NULL and hPipeStderrRd is NULL, stdout/stderr are merged
    HANDLE hPipeStdoutRd = NULL, hPipeStderrRd = NULL;

    // write handles for redirected output from the process
    // Both are NULL if not redirecting the process' stdout/stderr
    // If hStdoutRedirTarget is non-NULL and hStderrRedirTarget is NULL, stdout/stderr are merged
    HANDLE hStdoutRedirTarget = NULL, hStderrRedirTarget = NULL;

    // Handles to the threads monitoring the pipes for redirected stdout/stderr
    HANDLE hThread_StdoutMonitor = NULL, hThread_StderrMonitor = NULL;

    // ------------------------------------------------------------------------------------------

    /// <summary>
    /// Release all resources
    /// </summary>
    void Uninit();

    /// <summary>
    /// Explicit constructor
    /// </summary>
    ProcessInfo_t() = default;

    /// <summary>
    /// Explicit destructor
    /// </summary>
    ~ProcessInfo_t()
    {
        Uninit();
    }

private:
    // Because of the HANDLE members, do not allow copy or assignment
    ProcessInfo_t(const ProcessInfo_t&) = delete;
    ProcessInfo_t& operator = (const ProcessInfo_t&) = delete;
};

/// <summary>
/// Combine session- and process-specific information into a single structure
/// </summary>
struct SessionProcessInfo_t
{
    SessionInfo_t session;
    ProcessInfo_t process;

public:
    // Explicitly declare ctor/dtor
    SessionProcessInfo_t() = default;
    ~SessionProcessInfo_t() = default;

private:
    // No copy/assignment of these objects
    SessionProcessInfo_t(const SessionProcessInfo_t&) = delete;
    SessionProcessInfo_t& operator = (const SessionProcessInfo_t&) = delete;
};


/// <summary>
/// Reference-counted instances
/// </summary>
typedef std::shared_ptr<SessionProcessInfo_t> ptrSessionProcessInfo_t;

/// <summary>
/// Create and return an instance of a ptrSessionProcessInfo_t that can be passed to CreateThread as
/// the LPVOID lpParameter parameter.
/// Addresses of automatic/stack-based variables are not safe to pass to CreateThread as they might no longer
/// be valid by the time the thread begins execution.
/// </summary>
/// <param name="pSPI">Input: source smart pointer</param>
inline ptrSessionProcessInfo_t* CreateCrossThreadpSPI(const ptrSessionProcessInfo_t& pSPI)
{
    // Create a heap-allocated instance of the smart pointer and point it to the target object.
    ptrSessionProcessInfo_t* ppNewPointer = new ptrSessionProcessInfo_t;
    *ppNewPointer = pSPI;
    return ppNewPointer;
}

/// <summary>
/// For use within a CreateThread target: consume and deallocate a ptrSessionProcessInfo_t that
/// was allocated by CreateCrossThreadpSPI.
/// Points the pSPI smart pointer to the target object;
/// deallocates the heap-allocated object pointed to by lpvThreadParameter.
/// </summary>
/// <param name="pSPI">Reference: local smart pointer instance to point to the target object</param>
/// <param name="lpvThreadParameter">Address returned by CreateCrossThreadpSPI and had been passed as lpParameter to CreateThread</param>
inline void ConsumeCrossThreadpSPI(ptrSessionProcessInfo_t& pSPI, LPVOID lpvThreadParameter)
{
    pSPI = *(ptrSessionProcessInfo_t*)lpvThreadParameter;
    delete (ptrSessionProcessInfo_t*)lpvThreadParameter;
}

/// <summary>
/// Vector of reference-counted pointers to session/process-specific information
/// </summary>
typedef std::vector<ptrSessionProcessInfo_t> vecSessionProcessInfo_t;


/// <summary>
/// Class to manage processes started in users' sessions
/// </summary>
class ProcessManager_t
{
public:
    // Constructor - not much to do
    ProcessManager_t() { }
    // Destructor - release acquired resources
    ~ProcessManager_t() { Clear(); }

    // ------------------------------------------------------------------------------------------

    /// <summary>
    /// Create a new SessionProcessInfo_t object, add it to the collection, and return a pointer to it
    /// </summary>
    /// <returns>Pointer to a new SessionProcessInfo_t object in the managed collection</returns>
    ptrSessionProcessInfo_t New()
    {
        ptrSessionProcessInfo_t pNew = std::make_shared<SessionProcessInfo_t>();
        m_processes.push_back(pNew);
        return pNew;
    }

    void Clear();

    // ------------------------------------------------------------------------------------------

    /// <summary>
    /// Return a const iterator to the collection of allocated SessionProcessInfo_t instances
    /// </summary>
    vecSessionProcessInfo_t::const_iterator ConstIter() const
    {
        return m_processes.begin();
    }

    /// <summary>
    /// Return a non-const iterator to the collection of allocated SessionProcessInfo_t instances
    /// </summary>
    vecSessionProcessInfo_t::iterator Iter()
    {
        return m_processes.begin();
    }

    /// <summary>
    /// Indicate whether the const iterator is at the end of the collection
    /// </summary>
    /// <returns>true if at end; false otherwise</returns>
    bool IterAtEnd(vecSessionProcessInfo_t::const_iterator& iter) const
    {
        return iter == m_processes.end();
    }

    /// <summary>
    /// Indicate whether the non-const iterator is at the end of the collection
    /// </summary>
    /// <returns>true if at end; false otherwise</returns>
    bool IterAtEnd(vecSessionProcessInfo_t::iterator& iter)
    {
        return iter == m_processes.end();
    }

    // ------------------------------------------------------------------------------------------

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
    bool WaitForAProcessToExit(DWORD dwTimeout, DWORD& nRunningProcesses, DWORD& nNowRunning, vecSessionProcessInfo_t& vExitedProcesses);

    // ------------------------------------------------------------------------------------------

    /// <summary>
    /// For use when monitoring timeout expires: stop all threads monitoring pipes for redirected output, 
    /// and optionally terminate still-running processes.
    /// </summary>
    /// <param name="bTerminateProcesses">Input: whether to terminate still-running processes</param>
    void StopAllRedirectionMonitors(bool bTerminateProcesses);

    // ------------------------------------------------------------------------------------------

    /// <summary>
    /// Wait for all threads monitoring redirected output to exit
    /// </summary>
    void WaitForRedirectionMonitors();

    // ------------------------------------------------------------------------------------------

private:
    // Vector of pointers to allocated session/process info structures.
    vecSessionProcessInfo_t m_processes;

private:
    // Copy constructor and assignment operator not implemented
    ProcessManager_t(const ProcessManager_t&) = delete;
    ProcessManager_t& operator = (const ProcessManager_t&) = delete;
};

