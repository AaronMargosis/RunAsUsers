// Miscellaneous utility functions

#pragma once

#include <Windows.h>
#include <WtsApi32.h>
#include <locale>

/// <summary>
/// Base64-encode a string
/// </summary>
/// <param name="sInput">Input: string to encode</param>
/// <param name="sOutput">Output: encoded string</param>
/// <returns>true if successful; false otherwise</returns>
bool Base64Encode(const std::wstring& sInput, std::wstring& sOutput);

/// <summary>
/// Get the current system time as a 64-bit value representing the number of 100-nanosecond intervals 
/// since January 1, 1601 (UTC)..
/// </summary>
/// <param name="li">Output: current system time as a ULARGE_INTEGER</param>
inline void GetSystemTimeAsULargeinteger(ULARGE_INTEGER& li)
{
    // FILETIME is a pair of DWORDS - a high part and a low part, a 64-bit value representing the 
    // number of 100-nanosecond intervals since January 1, 1601 (UTC).
    // ULARGE_INTEGER is a union of the same thing along with a ULONGLONG, which simplifies math.
    // Use GetSystemTimeAsFileTime rather than GetSystemTimeAsPreciseFileTime to continue to support Win7/WS2008R2
    FILETIME ft = { 0 };
    GetSystemTimeAsFileTime(&ft);
    li.HighPart = ft.dwHighDateTime;
    li.LowPart = ft.dwLowDateTime;
}

/// <summary>
/// The number of milliseconds since the specified time (up to a max of 4,294,967,295 milliseconds -- ~49.7 days)
/// </summary>
/// <param name="liStart">Input: the start time to diff the current time against</param>
/// <returns>Number of milliseconds between current time and that start time (up to ~4 billion milliseconds)</returns>
inline DWORD MillisecondsSince(const ULARGE_INTEGER& liStart)
{
    // Contains a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (UTC).
    ULARGE_INTEGER liNow = { 0 };
    GetSystemTimeAsULargeinteger(liNow);
    ULARGE_INTEGER liDiff = { 0 };
    // (liNow.QuadPart - liStart.QuadPart) is the number of 100-nanosecond intervals since start time.
    // Divide by 10,000 to get the number of milliseconds.
    // liDiff.HighPart will be zero unless the time difference is over 4 billion milliseconds (beyond the
    // limits of this function's intended use).
    liDiff.QuadPart = (liNow.QuadPart - liStart.QuadPart) / (10 * 1000);
    return liDiff.LowPart;
}

/// <summary>
/// Convert WTS connection state enum into corresponding string
/// </summary>
inline const wchar_t* WtsConnectStateToWSZ(WTS_CONNECTSTATE_CLASS state)
{
    switch (state)
    {
    case WTSActive: return L"WTSActive";
    case WTSConnected: return L"WTSConnected";
    case WTSConnectQuery: return L"WTSConnectQuery";
    case WTSShadow: return L"WTSShadow";
    case WTSDisconnected: return L"WTSDisconnected";
    case WTSIdle: return L"WTSIdle";
    case WTSListen: return L"WTSListen";
    case WTSReset: return L"WTSReset";
    case WTSDown: return L"WTSDown";
    case WTSInit: return L"WTSInit";
    default: return L"[[[Unrecognized]]]";
    }
}


/// <summary>
/// Returns true if current system is Windows 7 or Windows Server 2008R2 (some things don't work correctly on that platform)
/// </summary>
bool IsWin7orWS2008R2();

/// <summary>
/// Convert WTS session-state (lock/unlock) info into corresponding string.
/// Note that Win7/WS2008R2 reports it incorrectly. Docs say they report the opposite of what they should; my experience is they always reported WTS_SESSIONSTATE_LOCK
/// </summary>
inline const wchar_t* WtsFlagsToWSZ(LONG wtsFlags)
{
    switch (wtsFlags)
    {
    case WTS_SESSIONSTATE_LOCK:
        return L"Locked";
    case WTS_SESSIONSTATE_UNLOCK:
        return L"Unlocked";
    default:
        return L"Unknown";
    }
}

