#pragma once
#include <Windows.h>
#include "CSid.h"

/// <summary>
/// A structure containing several attributes of a token.
/// </summary>
struct TokenInfo_t
{
	CSid sid;
	LUID logonSession;
	DWORD integrityLevel;
	//DWORD dwSessionId;

	/// <summary>
	/// Returns the integrity level value as a string
	/// </summary>
	std::wstring IntegrityLevelName() const;
};

class Token
{
public:
	//static bool GetTokenUser(HANDLE hToken, CSid& sid, std::wstring& sError);
	//static bool GetTokenLogonSession(HANDLE hToken, LUID logonSession, std::wstring& sError);
	//static bool GetTokenIntegrityLevel(HANDLE hToken, DWORD& dwIntegrityLevel, std::wstring& sError);
	//static bool GetTokenWtsSessionId(HANDLE hToken, DWORD& dwSessionId, std::wstring& sError);

	/// <summary>
	/// Retrieve user SID, logon session, and integrity level from the input token
	/// </summary>
	/// <param name="hToken">Input: token to inspect</param>
	/// <param name="tokenInfo">Output: structure containing token information</param>
	/// <param name="sErrorInfo">Output: error information on failure</param>
	/// <returns>true if successful, false otherwise</returns>
	static bool GetTokenInfo(HANDLE hToken, TokenInfo_t& tokenInfo, std::wstring& sErrorInfo);

	/// <summary>
	/// Get UAC-linked token, if present.
	/// Caller is responsible for closing the returned handle when done.
	/// </summary>
	/// <param name="hToken">Input token</param>
	/// <param name="hLinkedToken">UAC-linked token associated with input token, if present. NULL otherwise.</param>
	/// <returns>true if the input token has a UAC-linked token associated with it and returned as hLinkedToken.</returns>
	static bool GetLinkedToken(HANDLE hToken, HANDLE& hLinkedToken);

	/// <summary>
	/// If the supplied token is a UAC-limited token, get the elevated linked token associated with it and
	/// return it through the same parameter, closing the original token handle. (The caller is responsible for
	/// closing whatever token is returned through hToken.
	/// </summary>
	/// <param name="hToken">Input/output</param>
	/// <returns>Returns true if the token was swapped.
	/// Returns false if the input token has no linked token or is already the highest of the linked pair.</returns>
	static bool GetHighestToken(HANDLE& hToken);

public:
	Token() = delete;
	~Token() = delete;
	Token(const Token&) = delete;
	Token& operator = (const Token&) = delete;
};

