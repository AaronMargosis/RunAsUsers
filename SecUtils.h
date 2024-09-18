// Security utility functions
#pragma once

///// <summary>
///// Returns true if this process is running as SYSTEM (S-1-5-18).
///// </summary>
///// <returns>true if this process is running as SYSTEM; false otherwise.</returns>
//bool AmIRunningAsSystem();

///// <summary>
///// Creates a GUID and returns it in string form.
///// </summary>
///// <returns>GUID in string format</returns>
//std::wstring CreateNewGuidString();

///// <summary>
///// Get UAC-linked token, if present.
///// Caller is responsible for closing the returned handle when done.
///// </summary>
///// <param name="hToken">Input token</param>
///// <param name="hLinkedToken">UAC-linked token associated with input token, if present. NULL otherwise.</param>
///// <returns>true if the input token has a UAC-linked token associated with it and returned as hLinkedToken.</returns>
//bool GetLinkedToken(HANDLE hToken, HANDLE& hLinkedToken);
//
///// <summary>
///// If the supplied token is a UAC-limited token, get the elevated linked token associated with it and
///// return it through the same parameter, closing the original token handle. (The caller is responsible for
///// closing whatever token is returned through hToken.
///// </summary>
///// <param name="hToken">Input/output</param>
///// <returns>Returns true if the token was swapped.
///// Returns false if the input token has no linked token or is already the highest of the linked pair.</returns>
//bool GetHighestToken(HANDLE& hToken);

///// <summary>
///// Converts Win32 SID to std::wstring
///// </summary>
///// <param name="pSid">Input: SID to convert</param>
///// <param name="sSid">Output: SID in string form</param>
///// <returns>true if successful</returns>
//bool SidToString(PSID pSid, std::wstring& sSid);


