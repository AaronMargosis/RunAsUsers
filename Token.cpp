#include "Token.h"
#include "SysErrorMessage.h"
#include <sstream>

/// <summary>
/// Retrieve user SID, logon session, and integrity level from the input token
/// </summary>
/// <param name="hToken">Input: token to inspect</param>
/// <param name="tokenInfo">Output: structure containing token information</param>
/// <param name="sErrorInfo">Output: error information on failure</param>
/// <returns>true if successful, false otherwise</returns>
// static
bool Token::GetTokenInfo(HANDLE hToken, TokenInfo_t& tokenInfo, std::wstring& sErrorInfo)
{
	BYTE pBuffer[1024] = { 0 };
	DWORD dwLen = sizeof(pBuffer), dwReturnLength = 0;
	std::wstringstream strErrorInfo;

	if (GetTokenInformation(hToken, TokenUser, pBuffer, dwLen, &dwReturnLength))
	{
		PTOKEN_USER pTokInfo = (PTOKEN_USER)pBuffer;
		tokenInfo.sid = CSid(pTokInfo->User.Sid);
	}
	else
	{
		strErrorInfo << SysErrorMessageWithCode();
	}

	if (GetTokenInformation(hToken, TokenStatistics, pBuffer, dwLen, &dwReturnLength))
	{
		PTOKEN_STATISTICS pTokInfo = (PTOKEN_STATISTICS)pBuffer;
		tokenInfo.logonSession = pTokInfo->AuthenticationId;
	}
	else
	{
		strErrorInfo << SysErrorMessageWithCode();
	}

	if (GetTokenInformation(hToken, TokenIntegrityLevel, pBuffer, dwLen, &dwReturnLength))
	{
		PTOKEN_MANDATORY_LABEL pTokInfo = (PTOKEN_MANDATORY_LABEL)pBuffer;
		tokenInfo.integrityLevel = *GetSidSubAuthority(pTokInfo->Label.Sid, (DWORD)(UCHAR)(*GetSidSubAuthorityCount(pTokInfo->Label.Sid) - 1));
	}
	else
	{
		strErrorInfo << SysErrorMessageWithCode();
	}

	sErrorInfo = strErrorInfo.str();
	return true;
}

/// <summary>
/// Get UAC-linked token, if present.
/// Caller is responsible for closing the returned handle when done.
/// </summary>
/// <param name="hToken">Input token</param>
/// <param name="hLinkedToken">UAC-linked token associated with input token, if present. NULL otherwise.</param>
/// <returns>true if the input token has a UAC-linked token associated with it and returned as hLinkedToken.</returns>
// static
bool Token::GetLinkedToken(HANDLE hToken, HANDLE& hLinkedToken)
{
	// Initialize output variable.
	hLinkedToken = NULL;
	DWORD dwLength = sizeof(TOKEN_LINKED_TOKEN);
	TOKEN_LINKED_TOKEN linkedToken = { 0 };
	if (GetTokenInformation(hToken, TokenLinkedToken, &linkedToken, dwLength, &dwLength))
	{
		hLinkedToken = linkedToken.LinkedToken;
		return true;
	}
	else
		return false;
}

/// <summary>
/// If the supplied token is a UAC-limited token, get the elevated linked token associated with it and
/// return it through the same parameter, closing the original token handle. (The caller is responsible for
/// closing whatever token is returned through hToken.
/// </summary>
/// <param name="hToken">Input/output</param>
/// <returns>Returns true if the token was swapped.
/// Returns false if the input token has no linked token or is already the highest of the linked pair.</returns>
// static
bool Token::GetHighestToken(HANDLE& hToken)
{
	DWORD dwLength = sizeof(TOKEN_ELEVATION_TYPE);
	TOKEN_ELEVATION_TYPE elevType;
	// Determine the input token's elevation type
	if (GetTokenInformation(hToken, TokenElevationType, &elevType, dwLength, &dwLength))
	{
		// If the input token is the limited one of a pair, get the other one.
		if (TokenElevationTypeLimited == elevType)
		{
			HANDLE hLinkedToken = NULL;
			if (GetLinkedToken(hToken, hLinkedToken))
			{
				// Close the original token and replace it with the linked one.
				CloseHandle(hToken);
				hToken = hLinkedToken;
				return true;
			}
		}
	}
	return false;
}


std::wstring TokenInfo_t::IntegrityLevelName() const
{
	std::wstringstream strIL;

	switch (integrityLevel)
	{
	case SECURITY_MANDATORY_UNTRUSTED_RID:
		strIL << L"Untrusted";
		break;
	case SECURITY_MANDATORY_LOW_RID:
		strIL << L"Low";
		break;
	case SECURITY_MANDATORY_MEDIUM_RID:
		strIL << L"Medium";
		break;
	case SECURITY_MANDATORY_MEDIUM_PLUS_RID:
		strIL << L"MediumPlus";
		break;
	case SECURITY_MANDATORY_HIGH_RID:
		strIL << L"High";
		break;
	case SECURITY_MANDATORY_SYSTEM_RID:
		strIL << L"System";
		break;
	case SECURITY_MANDATORY_PROTECTED_PROCESS_RID:
		strIL << L"ProtectedProcess";
		break;
	default:
		strIL << integrityLevel;
		if (integrityLevel < SECURITY_MANDATORY_UNTRUSTED_RID)
			strIL << L" < Untrusted";
		else if (integrityLevel < SECURITY_MANDATORY_LOW_RID)
			strIL << L" < Low";
		else if (integrityLevel < SECURITY_MANDATORY_MEDIUM_RID)
			strIL << L" < Medium";
		else if (integrityLevel < SECURITY_MANDATORY_MEDIUM_PLUS_RID)
			strIL << L" < MediumPlus";
		else if (integrityLevel < SECURITY_MANDATORY_HIGH_RID)
			strIL << L" < High";
		else if (integrityLevel < SECURITY_MANDATORY_SYSTEM_RID)
			strIL << L" < System";
		else if (integrityLevel < SECURITY_MANDATORY_PROTECTED_PROCESS_RID)
			strIL << L" < ProtectedProcess";
		else
			strIL << L" > ProtectedProcess";
		break;
	}

	return strIL.str();
}
