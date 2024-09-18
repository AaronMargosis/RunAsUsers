// String utilities

#include <Windows.h>
#include <sstream>
#include <locale>

#include "StringUtils.h"

/// <summary>
/// Similar to .NET's string split method, returns a vector of substrings of the input string based
/// on the supplied delimiter.
/// </summary>
/// <param name="strInput">Input: string from which to return substrings</param>
/// <param name="delim">Input: delimiter character to separate substrings</param>
/// <param name="elems">Output: vector of substrings</param>
void SplitStringToVector(const std::wstring& strInput, wchar_t delim, std::vector<std::wstring>& elems)
{
	elems.clear();
	// If input string is zero length, return a zero-length vector.
	if (strInput.length() == 0)
		return;
	std::wstringstream ss(strInput);
	std::wstring item;
	// Get everything up to EOF. Problem with testing operator bool() on getline's return value is that
	// if the string ends with a delimiter, the last field ends up getting dropped. This technique
	// fixes that.
	do {
		std::getline(ss, item, delim);
		elems.push_back(item);
	} while (!ss.eof());
}

// ------------------------------------------------------------------------------------------
/// <summary>
/// Convert a wstring in place to locale-sensitive upper-case
/// </summary>
/// <param name="str"></param>
/// <returns></returns>
std::wstring& WString_To_Upper(std::wstring& str)
{
	// Note my test results against all wchar_t values 0x0000 to 0xFFFF:
	// ::toupper(c) and ::tolower(c) each change 26 characters out of 65536
	// std::toupper(c, loc) and std::tolower(c, loc) where loc is initialized with "" each change
	//    973 characters out of 65536.
	// std::toupper(c, loc) and std::tolower(c, loc) where loc is std::locale::empty change
	//    943 and 942 characters, respectively.
	std::locale loc("");
	size_t len = str.length();
	for (size_t ix = 0; ix < len; ++ix)
	{
		str[ix] = std::toupper(str[ix], loc);
	}
	return str;
}

// ----------------------------------------------------------------------------------------------------
// Date/time-related string manipulation

/// <summary>
/// Convert input system time structure to an alpha-sortable date/time string, optionally including 
/// milliseconds, and optionally including only characters that are valid in directory and file names.
/// </summary>
/// <param name="st">Input: SYSTEMTIME structure representing the date/time to convert to a string</param>
/// <param name="bIncludeMilliseconds">Input: true to include milliseconds, false otherwise</param>
/// <param name="bForFileSystem">Input: true to limit to file-object-valid characters</param>
/// <returns>Timestamp string with a format like yyyy-MM-dd HH:mm:ss.fff</returns>
std::wstring SystemTimeToWString(const SYSTEMTIME& st, bool bIncludeMilliseconds, bool bForFileSystem)
{
	wchar_t szTimestamp[32];
	const size_t bufferSize = sizeof(szTimestamp) / sizeof(szTimestamp[0]);
	if (bIncludeMilliseconds)
	{
		swprintf(szTimestamp, bufferSize, 
			bForFileSystem ? L"%04d%02d%02d_%02d%02d%02d_%03d" : L"%04d-%02d-%02d %02d:%02d:%02d.%03d",
			st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
	}
	else
	{
		swprintf(szTimestamp, bufferSize, 
			bForFileSystem ? L"%04d%02d%02d_%02d%02d%02d" : L"%04d-%02d-%02d %02d:%02d:%02d",
			st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
	}
	return szTimestamp;
}

/// <summary>
/// Convert input filetime structure to an alpha-sortable date/time string, optionally including
/// milliseconds and optionally including only characters that are valid in directory and file names.
/// If the filetime value is zero, returns an alternate string value (rather than converting to 1601-01-01).
/// </summary>
/// <param name="ft">Input: FILETIME structure representing the date/time to convert to a string</param>
/// <param name="bIncludeMilliseconds">Input: true to include milliseconds, false otherwise</param>
/// <param name="szIfZero">Input (optional): the string to return if the ft is zero.</param>
/// <returns>Timestamp string with a format like yyyy-MM-dd HH:mm:ss.fff, or alternate string value</returns>
std::wstring FileTimeToWString(const FILETIME& ft, bool bIncludeMilliseconds, const wchar_t* szIfZero)
{
	if (0 == ft.dwHighDateTime && 0 == ft.dwLowDateTime)
		return szIfZero ? szIfZero : L"";

	SYSTEMTIME st;
	FileTimeToSystemTime(&ft, &st);
	return SystemTimeToWString(st, bIncludeMilliseconds, false);
}

/// <summary>
/// Converts input LARGE_INTEGER to an alpha-sortable date/time string, where the input value represents
/// the number of 100-nanosecond intervals since January 1, 1601 (UTC).
/// </summary>
/// <param name="l">Input: Number of 100-nanosecond intervals since 1/1/1601</param>
/// <param name="bIncludeMilliseconds">Input: true to include milliseconds, false otherwise</param>
/// <param name="szIfZero">Input (optional): the string to return if the ft is zero.</param>
/// <returns>Timestamp string with a format like yyyy-MM-dd HH:mm:ss.fff, or alternate string value</returns>
std::wstring LargeIntegerToDateTimeString(const LARGE_INTEGER& l, bool bIncludeMilliseconds, const wchar_t* szIfZero)
{
	FILETIME ft;
	ft.dwHighDateTime = l.HighPart;
	ft.dwLowDateTime = l.LowPart;
	return FileTimeToWString(ft, bIncludeMilliseconds, szIfZero);
}

/// <summary>
/// Creates and returns an alpha-sortable timestamp string from the current time, optionally including milliseconds
/// </summary>
/// <param name="bIncludeMilliseconds">Input: whether to incorporate milliseconds in the output</param>
/// <returns>Alpha-sortable timestamp string</returns>
std::wstring TimestampUTC(bool bIncludeMilliseconds /*= false*/)
{
	//wchar_t szTimestamp[32];
	SYSTEMTIME st;
	GetSystemTime(&st);
	return SystemTimeToWString(st, bIncludeMilliseconds, false);
}

/// <summary>
/// Create and returns alpha-sortable timestamp string from the current time, optionally including milliseconds,
/// using only characters that are valid in filenames.
/// </summary>
/// <param name="bIncludeMilliseconds">Input: whether to incorporate milliseconds in the output</param>
/// <returns>Alpha-sortable timestamp string</returns>
std::wstring TimestampUTCforFilepath(bool bIncludeMilliseconds /*= false*/)
{
	//wchar_t szTimestamp[32];
	SYSTEMTIME st;
	GetSystemTime(&st);
	return SystemTimeToWString(st, bIncludeMilliseconds, true);
}

// ------------------------------------------------------------------------------------------
// File system path string manipulation

/// <summary>
/// Splits the input file path into three parts: directory, filename (without extension), extension.
/// If filename ends with a dot, it's included as part of the filenameNoExt.
/// </summary>
bool SplitFilePath(const std::wstring& sFullPath, std::wstring& sDirectory, std::wstring& sFilenameNoExt, std::wstring& sExtension)
{
	sDirectory.clear();
	sFilenameNoExt.clear();
	sExtension.clear();

	std::wstring sFilename;

	// Look for last forward slash or backslash
	size_t ixLastPathSep = sFullPath.find_last_of(L"/\\");
	if (std::wstring::npos == ixLastPathSep)
	{
		// No path separators in the path - the input string is just a file name
		sFilename = sFullPath;
	}
	else
	{
		// Filename is the string following the last path separator
		sDirectory = sFullPath.substr(0, ixLastPathSep);
		sFilename = sFullPath.substr(ixLastPathSep + 1);
	}

	// Look for last dot in file name
	size_t ixLastDot = sFilename.find_last_of(L'.');
	if (std::wstring::npos == ixLastDot)
	{
		// No dots in file name
		sFilenameNoExt = sFilename;
	}
	else
	{
		sFilenameNoExt = sFilename.substr(0, ixLastDot);
		sExtension = sFilename.substr(ixLastDot + 1); // text after last dot
		// If filename ends with dot, add it back to the filename
		if (0 == sExtension.length())
		{
			sFilenameNoExt.append(L".");
		}
	}

	return true;
}

/// <summary>
/// Returns the parent directory of the input file specification, or an empty string if there is no parent directory.
/// Handles UNC paths correctly (highest-level directory includes server and share name).
/// Forward slash and backslash are both valid path separators.
/// Note that parent directory of "C:\Subdir" will be returned as "C:", not "C:\" with the trailing slash.
/// </summary>
/// <param name="sFilePath">Input file specification (can be a path to a file or a directory)</param>
/// <returns>Parent directory, or empty string if no parent directory</returns>
std::wstring GetDirectoryNameFromFilePath(const std::wstring& sFilePath)
{
	// Accept both forward slash or backslash as path separators
	const wchar_t* szPathSepChars = L"/\\";

	// Find last path separator in the input. If none found, return an empty string.
	size_t ixLastPathSep = sFilePath.find_last_of(szPathSepChars);
	if (std::wstring::npos == ixLastPathSep)
	{
		return std::wstring();
	}

	// UNC path if the input starts with \\ or //. For UNC, shortest valid parent directory name includes the server and the share specification;
	// i.e., of the form \\server\share.
	if (StartsWith(sFilePath, L"\\\\") || StartsWith(sFilePath, L"//"))
	{
		// First, find the first path separator after the server name:
		size_t ixPathSep = sFilePath.find_first_of(szPathSepChars, 2);
		// If that's found, then find the first path separator after the share name:
		if (std::wstring::npos != ixPathSep)
			ixPathSep = sFilePath.find_first_of(szPathSepChars, ixPathSep + 1);
		// If there aren't enough path separators, or the last one is the last character in the string, there's no parent directory; return empty
		if (std::wstring::npos == ixPathSep || sFilePath.length() - 1 == ixPathSep)
			return std::wstring();
	}
	else
	{
		// Not a UNC path.
		// If the first path separator is the last character in the string (e.g., "C:\"), there's no parent directory; return empty.
		// (Note that because of the ixLastPathSep check earlier, ixFirstPathSep is guaranteed not to be npos, and sFilePath is not an empty string.)
		size_t ixFirstPathSep = sFilePath.find_first_of(szPathSepChars);
		if (sFilePath.length() - 1 == ixFirstPathSep)
			return std::wstring();
	}

	// All verifications check out. Return the string up to the last path separator.
	return sFilePath.substr(0, ixLastPathSep);
}
