#pragma once

#include <fstream>
#include <string>

/// <summary>
/// Ensure that output stream produces UTF-8 with optional BOM
/// </summary>
/// <param name="stream">Output stream to imbue</param>
/// <param name="bGenerateHeader">true to generate BOM at start of output sequence, false not to</param>
void ImbueStreamUtf8(std::wostream& stream, bool bGenerateHeader = true);

/// <summary>
/// Creates a output file stream for UTF-8 output with BOM.
/// </summary>
/// <param name="szFilename">Input: name of output file</param>
/// <param name="fOutput">Output: resulting wofstream object</param>
/// <param name="bAppend">Input: true to append to file, false to overwrite (default)</param>
/// <returns>true on success, false otherwise</returns>
bool CreateFileOutput(const wchar_t* szFilename, std::wofstream& fOutput, bool bAppend = false);
