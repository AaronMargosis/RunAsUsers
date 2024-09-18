#include <Windows.h>
#include <WtsApi32.h>
#pragma comment(lib, "wtsapi32.lib")
#include "FileOutput.h"
#include "DbgOut.h"
#include "WofstreamManager.h"
#include "StringUtils.h"

// ------------------------------------------------------------------------------------------

/// <summary>
/// Global instance defined. You can create other instances each with their own targets enabled. 
/// </summary>
DbgOut_t dbgOut;

// Single instance of a managed collection of shareable std::wofstream instances.
// Heap-allocated when the first instance of DbgOut_InternalBufferImpl is created, and never deleted,
// to ensure that the destructors of global instances of DbgOut_InternalBufferImpl defined in arbitrary
// other compilation units always have a valid WofstreamManager_t to access.
WofstreamManager_t* DbgOut_InternalBufferImpl::st_pWofstreamMgr = nullptr;

// Singleton instances of critical section objects to serialize access to std::wcout and std::wcerr.
// Bool to indicate whether they've been initialized.
static CRITICAL_SECTION gb_critsecWCout, gb_critsecWCerr;
static bool gb_bCritSecsInitialized = false;

// ------------------------------------------------------------------------------------------

// By default, only debug stream is enabled.
DbgOut_InternalBufferImpl::DbgOut_InternalBufferImpl() :
	m_bWriteToDebugStream(true),
	m_bWriteToWCout(false),
	m_bWriteToWCerr(false),
	m_bWriteToWtsMsgBox(false),
	m_bWriteToFile(false),
	m_bAcquiredOutput(false),
	m_bPrependTimestamp(false)
{
	// First one in instantiates a heap-allocated instance of the WofstreamManager_t for which the
	// destructor never gets called. This ensures that global instances of DbgOut_InternalBufferImpl
	// defined in other compilation units always have a valid WofstreamManager_t to use when their
	// destructors are invoked
	if (nullptr == st_pWofstreamMgr)
	{
		st_pWofstreamMgr = new WofstreamManager_t;
	}

	// Initialize this instance's synchronization mechanisms.
	InitializeCriticalSection(&m_critsecConfig);
	InitializeCriticalSection(&m_critsecOutput);

	// and initialize the global wcout/wcerr critical sections if not already done
	if (!gb_bCritSecsInitialized)
	{
		gb_bCritSecsInitialized = true;
		InitializeCriticalSection(&gb_critsecWCout);
		InitializeCriticalSection(&gb_critsecWCerr);
	}

}

DbgOut_InternalBufferImpl::~DbgOut_InternalBufferImpl()
{
	// Serialize configuration changes
	EnterCriticalSection(&m_critsecConfig);
	// Push out anything that was still left in the buffer
	sync();
	// Release any file that this instance happens to have open
	releaseFile();
	LeaveCriticalSection(&m_critsecConfig);

	// Done with these now
	DeleteCriticalSection(&m_critsecConfig);
	DeleteCriticalSection(&m_critsecOutput);
}

void DbgOut_InternalBufferImpl::WriteToDebugStream(bool bWriteToDebugStream)
{
	// Serialize configuration changes
	EnterCriticalSection(&m_critsecConfig);
	m_bWriteToDebugStream = bWriteToDebugStream;
	LeaveCriticalSection(&m_critsecConfig);
}

void DbgOut_InternalBufferImpl::WriteToWCout(bool bWriteToWCout)
{
	// Serialize configuration changes
	EnterCriticalSection(&m_critsecConfig);
	m_bWriteToWCout = bWriteToWCout;
	LeaveCriticalSection(&m_critsecConfig);
}

void DbgOut_InternalBufferImpl::WriteToWCerr(bool bWriteToWCerr)
{
	// Serialize configuration changes
	EnterCriticalSection(&m_critsecConfig);
	m_bWriteToWCerr = bWriteToWCerr;
	LeaveCriticalSection(&m_critsecConfig);
}

void DbgOut_InternalBufferImpl::WriteToWtsMsgBox(bool bWriteToWtsMsgBox)
{
	// Serialize configuration changes
	EnterCriticalSection(&m_critsecConfig);
	m_bWriteToWtsMsgBox = bWriteToWtsMsgBox;
	LeaveCriticalSection(&m_critsecConfig);
}

bool DbgOut_InternalBufferImpl::WriteToFile(const wchar_t* szFilename, bool bAppend /*= false*/, uint64_t uSizeThreshold /* = 0*/)
{
	bool retval = true;
	// Serialize configuration changes
	EnterCriticalSection(&m_critsecConfig);
	// try/catch block to ensure that the LeaveCriticalSection call gets made.
	try
	{
		// Release any existing file if one is referenced.
		releaseFile();
		if (szFilename && *szFilename)
		{
			// Try to get a pointer to a (possibly shared) std::wofstream instance.
			if (st_pWofstreamMgr->GetWofstream(szFilename, &m_pStreamSync, bAppend, uSizeThreshold))
			{
				m_bWriteToFile = true;
				retval = true;
			}
			else
			{
				retval = false;
			}
		}
	}
	catch (...) {}
	LeaveCriticalSection(&m_critsecConfig);
	return retval;
}

void DbgOut_InternalBufferImpl::WriteToHANDLE(HANDLE handle)
{
	m_handle = handle;
}

void DbgOut_InternalBufferImpl::PrependTimestamp(bool bPrependTimestamp)
{
	// Serialize configuration changes
	EnterCriticalSection(&m_critsecConfig);
	m_bPrependTimestamp = bPrependTimestamp;
	LeaveCriticalSection(&m_critsecConfig);
}

// Acquire access to this instance's critical section.
void DbgOut_InternalBufferImpl::acquireOutput()
{
	EnterCriticalSection(&m_critsecOutput);
	// After acquiring, set the flag that can be checked on release.
	m_bAcquiredOutput = true;
}

void DbgOut_InternalBufferImpl::releaseOutput()
{
	// Don't do anything with the critical section if it wasn't acquired.
	//TODO: This works only if all threads are using the synch mechanism; consider putting this value in thread-local storage.
	if (m_bAcquiredOutput)
	{
		// Clear the flag while still holding the critical section.
		m_bAcquiredOutput = false;
		// If this thread called acquireOutput more than once, release them all.
		// Get the recursion count before beginning to release them.
		LONG releaseCount = m_critsecOutput.RecursionCount;
		for (LONG ix = 0; ix < releaseCount; ++ix)
		{
			LeaveCriticalSection(&m_critsecOutput);
		}
	}
}

// Local function that puts a message box on the desktop of all active WTS sessions.
// If it fails, it fails silently.
static void ToWtsMsgBox(const std::wstring& str)
{
	PWTS_SESSION_INFOW pSessionInfo = NULL;
	DWORD dwSessionCount = 0;
	BOOL ret = WTSEnumerateSessionsW(WTS_CURRENT_SERVER_HANDLE, 0, 1, &pSessionInfo, &dwSessionCount);
	if (!ret) return;

	wchar_t szDebugTitle[] = L"Debug Output";
#pragma warning(push)
#pragma warning(disable: 4267) // possible loss of data converting from size_t to DWORD
	DWORD dwTitleLength = wcslen(szDebugTitle) * sizeof(wchar_t);
	DWORD dwMessageLength = str.length() * sizeof(wchar_t);
#pragma warning(pop)
	DWORD dwResponse = 0;
	for (DWORD ix = 0; ix < dwSessionCount; ++ix)
	{
		if (WTSActive == pSessionInfo[ix].State)
		{
			WTSSendMessageW(
				WTS_CURRENT_SERVER_HANDLE,
				pSessionInfo[ix].SessionId,
				szDebugTitle,
				dwTitleLength,
				(LPWSTR)str.c_str(),
				dwMessageLength,
				MB_OK | MB_ICONINFORMATION,
				0,
				&dwResponse,
				FALSE);
		}
	}
	WTSFreeMemory(pSessionInfo);
}

// Override of wstringbuf member function that triggers write to destination(s)
int DbgOut_InternalBufferImpl::sync()
{
	// Don't do anything if buffer is empty
	if (str().length() > 0)
	{
		// Catch exceptions to ensure that releaseOutput() is executed
		try
		{
			std::wstring sOutput;
			if (!m_bPrependTimestamp)
			{
				sOutput = str();
			}
			else
			{
				// Build timestamp string:
				std::wstring sTimestamp = TimestampUTC(true);

				// Split the output string on LF (don't need to worry about CR, as we're going to add LF back in anyway)
				// Note that if the output ends with LF, the last element in lines will be a zero-length string.
				std::vector<std::wstring> lines;
				SplitStringToVector(str(), L'\n', lines);
				const size_t nLines = lines.size();

				// Build the real output as strOutput
				std::wstringstream strOutput;
				// Insert timestamp and line text
				for (size_t ixLine = 0; ixLine < nLines - 1; ++ixLine)
				{
					strOutput << sTimestamp << L": " << lines[ixLine] << L'\n';
				}
				// If the last element in lines is not a zero-length string, the original output did not end with LF.
				// If that is the case, add that text into the output buffer but with no trailing LF.
				if (lines[nLines - 1].length() > 0)
				{
					strOutput << sTimestamp << L": " << lines[nLines - 1];
				}
				sOutput = strOutput.str();
			}

			// debug stream is threadsafe
			if (m_bWriteToDebugStream)
			{
				OutputDebugStringW(sOutput.c_str());
			}
			if (m_bWriteToWCout)
			{
				// Serialize access to std::wcout
				EnterCriticalSection(&gb_critsecWCout);
				// Exception handling to ensure that the Leave API gets called
				try { std::wcout << sOutput << std::flush; } catch(...) {}
				LeaveCriticalSection(&gb_critsecWCout);
			}
			if (m_bWriteToWCerr)
			{
				// Serialize access to std::wcerr
				EnterCriticalSection(&gb_critsecWCerr);
				// Exception handling to ensure that the Leave API gets called
				try { std::wcerr << sOutput << std::flush; } catch(...) {}
				LeaveCriticalSection(&gb_critsecWCerr);
			}
			// WTS message box doesn't need serialization
			if (m_bWriteToWtsMsgBox)
			{
				ToWtsMsgBox(sOutput);
			}
			if (m_bWriteToFile && m_pStreamSync)
			{
				// Serialize access to this (possibly shared) std::wofstream
				EnterCriticalSection(&m_pStreamSync->m_critsec);
				// Exception handling to ensure that the Leave API gets called
				try { m_pStreamSync->m_fstream << sOutput << std::flush; } catch(...) {}
				// Also enforce the file size threshold
				try { m_pStreamSync->EnforceSizeThreshold(); } catch (...) {}
				LeaveCriticalSection(&m_pStreamSync->m_critsec);
			}
			if (m_handle)
			{
				// Not serializing, just raw writes.
				DWORD dwBytes = DWORD(sOutput.length() * sizeof(wchar_t));
				WriteFile(m_handle, sOutput.c_str(), dwBytes, &dwBytes, nullptr);
			}
		}
		catch (...) {}

		// Clear the buffer
		str(L"");

		// Release access to output buffer if acquired
		releaseOutput();
	}
	return 0;
}

// Internal: close/release a log file if one is open
void DbgOut_InternalBufferImpl::releaseFile()
{
	if (m_bWriteToFile)
	{
		if (m_pStreamSync)
		{
			// WofstreamManager_t is responsible for allocating/deallocating instances.
			// We don't delete it here. Just don't reference it anymore.
			st_pWofstreamMgr->ReleaseWofstream(m_pStreamSync);
			m_pStreamSync = nullptr;
		}
		m_bWriteToFile = false;
	}
}

// ------------------------------------------------------------------------------------------

/// <summary>
/// Calling this method acquires exclusive access to the object's output buffer until the caller
/// inserts std::endl or std::flush into it. It returns a reference to the current object so that
/// it can be used directly in the stream; e.g.,
///     dbgOut.locked() << L"Information: " << pvAddr << std::endl;
/// (Note that "exclusive access" works only if callers in other threads also use the .locked() accessor.)
/// </summary>
DbgOut_t& DbgOut_t::locked()
{
	m_buf.acquireOutput();
	return *this;
}

void DbgOut_t::WriteToDebugStream(bool bWriteToDebugStream)
{
	m_buf.WriteToDebugStream(bWriteToDebugStream);
}

void DbgOut_t::WriteToWCout(bool bWriteToWCout)
{
	m_buf.WriteToWCout(bWriteToWCout);
}

void DbgOut_t::WriteToWCerr(bool bWriteToWCerr)
{
	m_buf.WriteToWCerr(bWriteToWCerr);
}

void DbgOut_t::WriteToWtsMsgBox(bool bWriteToWtsMsgBox)
{
	m_buf.WriteToWtsMsgBox(bWriteToWtsMsgBox);
}

bool DbgOut_t::WriteToFile(const wchar_t* szFilename, bool bAppend /*= false*/, uint64_t uSizeThreshold /* = 0*/)
{
	return m_buf.WriteToFile(szFilename, bAppend, uSizeThreshold);
}

void DbgOut_t::WriteToHANDLE(HANDLE handle)
{
	m_buf.WriteToHANDLE(handle);
}

void DbgOut_t::PrependTimestamp(bool bPrependTimestamp)
{
	m_buf.PrependTimestamp(bPrependTimestamp);
}
