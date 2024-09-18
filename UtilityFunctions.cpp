// Miscellaneous utility functions

#include <Windows.h>
#include <wincrypt.h>
#pragma comment(lib, "crypt32.lib")
#include <string>
#include <sstream>
#include <vector>
#include "SysErrorMessage.h"
#include "StringUtils.h"
#include "UtilityFunctions.h"

// ----------------------------------------------------------------------------------------------------
/// <summary>
/// Base64-encode a string
/// </summary>
/// <param name="sInput">Input: string to encode</param>
/// <param name="sOutput">Output: encoded string</param>
/// <returns>true if successful; false otherwise</returns>
bool Base64Encode(const std::wstring& sInput, std::wstring& sOutput)
{
    bool retval = false;
    sOutput.clear();

    DWORD cchOutput = 0;
#pragma warning(push)
// When compiling for x64, this line triggers C4267 'initializing': conversion from 'size_t' to 'DWORD', possible loss of data
// No warning when building for x86.
#pragma warning(disable: 4267)
    DWORD cbBinary = sInput.length() * sizeof(sInput[0]);
#pragma warning(pop)
    // Determine size of buffer we need to allocate.
    BOOL ret = CryptBinaryToStringW((const BYTE*)sInput.c_str(), cbBinary, CRYPT_STRING_BASE64, nullptr, &cchOutput);
    DWORD dwLastErr = GetLastError();
    if (cchOutput > 0)
    {
        std::vector<wchar_t> vChars(cchOutput + 8);
        ret = CryptBinaryToStringW((const BYTE*)sInput.c_str(), cbBinary, CRYPT_STRING_BASE64, vChars.data(), &cchOutput);
        dwLastErr = GetLastError();
        if (ret)
        {
            sOutput = vChars.data();
            // Remove CRLF
            sOutput = replaceStringAll(sOutput, L"\r\n", L"");
            retval = true;
        }
    }

    return retval;
}

/// <summary>
/// Returns true if current system is Windows 7 or Windows Server 2008R2, any service pack.
/// (Reason for test is that some things don't work correctly on that platform)
/// </summary>
bool IsWin7orWS2008R2()
{
    // Windows 7 / WS2008R2, any service pack
    DWORDLONG const dwlConditionMask = 
        VerSetConditionMask(
        VerSetConditionMask(
            0, 
            VER_MAJORVERSION, VER_EQUAL),
            VER_MINORVERSION, VER_EQUAL);

	OSVERSIONINFOEXW osvi = { sizeof(osvi), 0, 0, 0, 0, {0}, 0, 0 };
    osvi.dwMajorVersion = HIBYTE(_WIN32_WINNT_WIN7);
    osvi.dwMinorVersion = LOBYTE(_WIN32_WINNT_WIN7);

    BOOL vviRet = VerifyVersionInfoW(
        &osvi, 
        VER_MAJORVERSION | VER_MINORVERSION, 
        dwlConditionMask);

    return (FALSE != vviRet);
}

