#pragma once

#include "ProcessManager.h"

/// <summary>
/// Sets up everything for redirecting a target process' stdout/stderr to a destination
/// </summary>
/// <param name="pSPI">ptrSessionProcessInfo_t for the process to monitor</param>
/// <param name="bRedirStd">Input: whether to redirect stdout/stderr</param>
/// <param name="bMergeStd">Input: whether to merge stderr into stdout</param>
/// <param name="sRedirStdDirectory">Input: directory in which to create output file(s) for redirection; if empty, then redirect to this process' stdout/stderr</param>
/// <returns></returns>
bool SetUpRedirection(
	ptrSessionProcessInfo_t& pSPI,
	bool bRedirStd,
	bool bMergeStd,
	const std::wstring& sRedirStdDirectory
);
