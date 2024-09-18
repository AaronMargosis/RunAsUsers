#include "FileOutput.h"
#include <locale>
#include <codecvt>
#include <Windows.h>

/// <summary>
/// Ensure that output stream produces UTF-8 with optional BOM
/// </summary>
/// <param name="stream">Output stream to imbue</param>
void ImbueStreamUtf8(std::wostream& stream, bool bGenerateHeader)
{
    // Ensure that stream output is UTF-8.
    // Note that the heap-allocated std::codecvt_utf8 will eventually be deleted by the std::locale
    // it's initializing, so we MUST NOT match that "new" with a "delete" here.
    if (bGenerateHeader)
    {
        std::locale loc(std::locale(), new std::codecvt_utf8<wchar_t, 0x10ffff, std::generate_header>);
        stream.imbue(loc);
    }
    else
    {
        std::locale loc(std::locale(), new std::codecvt_utf8<wchar_t, 0x10ffff>);
        stream.imbue(loc);
    }
}


/// <summary>
/// Creates a output file stream for UTF-8 output with BOM.
/// </summary>
/// <param name="szFilename">Input: name of output file</param>
/// <param name="fOutput">Output: resulting wofstream object</param>
/// <param name="bAppend">Input: true to append to file, false to overwrite (default)</param>
/// <returns>true on success, false otherwise</returns>
bool CreateFileOutput(const wchar_t* szFilename, std::wofstream & fOutput, bool bAppend /*= false*/)
{
    // If appending and the file already exists and is more than 0 bytes in length, do not generate the BOM header.
    // If it doesn't exist or is zero-length, append doesn't matter, so use that bool to determine whether to 
    // generate the BOM.
    if (bAppend)
    {
        WIN32_FILE_ATTRIBUTE_DATA data = { 0 };
        if (GetFileAttributesExW(szFilename, GetFileExInfoStandard, &data))
        {
            if (0 == data.nFileSizeHigh && 0 == data.nFileSizeLow)
            {
                bAppend = false;
            }
        }
        else
        {
            DWORD dwLastErr = GetLastError();
            if (ERROR_FILE_NOT_FOUND == dwLastErr)
            {
                bAppend = false;
            }
        }
    }
    fOutput.open(szFilename, (bAppend ? (std::ios_base::out | std::ios_base::app) : std::ios_base::out));
    if (fOutput.fail())
    {
        return false;
    }
    // Ensure that file output is UTF-8. BOM unless appending to a non-empty existing file.
    ImbueStreamUtf8(fOutput, !bAppend);
    return true;
}
