// Manages serialization of access to possibly-shared std::wofstream instances.
//
// If two instances of std::wofstream write to the same file, the results are "undefined" at best, and
// can cause memory access violations. This code enables multiple threads to enjoy serialized access to
// possibly-shared std::wofstream instances. Filenames are canonicalized so that the code can determine
// when multiple references use different names to access the same file (e.g., long vs. short names, 
// relative vs. absolute paths, upper vs. lower case).

#include <locale>
#include <sstream>
#include "WofstreamManager.h"
#include "FileOutput.h"
#include "StringUtils.h"


// ------------------------------------------------------------------------------------------
// WofstreamSync_t
// 

// Constructor
WofstreamSync_t::WofstreamSync_t()
	: m_uSizeThreshold(0)
{
	InitializeCriticalSection(&m_critsec);
}

// Destructor
WofstreamSync_t::~WofstreamSync_t()
{
	DeleteCriticalSection(&m_critsec);
	// Close file if still open
	if (m_fstream.is_open())
	{
		m_fstream.close();
	}
}

/// <summary>
/// Enforces the file's size threshold (if non-zero). If the size has reached or exceeded 
/// the threshold, closes the stream, renames the file with a timestamp in the file name, 
/// then opens a new stream with the original file name.
/// CALLER MUST HAVE ACQUIRED THE CRITICAL SECTION.
/// </summary>
void WofstreamSync_t::EnforceSizeThreshold()
{
	// 0 means no max size
	if (0 != m_uSizeThreshold)
	{
		// Make sure it's open before proceeding
		if (m_fstream.is_open())
		{
			// Get the file size and compare to the max.
			// If can't acquire file size for any reason, don't do anything
			WIN32_FILE_ATTRIBUTE_DATA data = { 0 };
			if (GetFileAttributesExW(m_sCanonicalizedNameCasePreserved.c_str(), GetFileExInfoStandard, &data))
			{
				// Get the file size in a form that we can compare to the threshold.
				ULARGE_INTEGER ul;
				ul.HighPart = data.nFileSizeHigh;
				ul.LowPart = data.nFileSizeLow;
				if (ul.QuadPart >= m_uSizeThreshold)
				{
					// Build the new file name
					std::wstring sDirectory, sFilenameNoExt, sExtension, sTimestamp;
					std::wstringstream strNewFilename;
					SplitFilePath(m_sCanonicalizedNameCasePreserved, sDirectory, sFilenameNoExt, sExtension);
					sTimestamp = TimestampUTCforFilepath(true);
					strNewFilename << sFilenameNoExt << L"_" << sTimestamp;
					if (sExtension.length() > 0)
					{
						strNewFilename << L"." << sExtension;
					}
					m_fstream.close();
					BOOL ret = MoveFileW(m_sCanonicalizedNameCasePreserved.c_str(), strNewFilename.str().c_str());
					if (!ret)
					{
						DWORD dwLastErr = GetLastError();
						std::wstringstream strError;
						strError << L"MoveFileW failed, error " << dwLastErr << std::endl
							<< L"Source:  " << m_sCanonicalizedNameCasePreserved << std::endl
							<< L"NewName: " << strNewFilename.str() << std::endl;
						OutputDebugStringW(strError.str().c_str());
					}
					// Should be new file, but if the rename didn't succeed, append to the old rather than overwrite.
					CreateFileOutput(m_sCanonicalizedNameCasePreserved.c_str(), m_fstream, true);
				}
			}
		}

	}
}

/// <summary>
/// Increase this object's reference count
/// </summary>
/// <returns>New reference count</returns>
size_t WofstreamSync_t::AddRef()
{
	EnterCriticalSection(&m_critsec);
	++m_refCount;
	LeaveCriticalSection(&m_critsec);
	return m_refCount;
}

/// <summary>
/// Decrement this object's reference count. If decremented to 0, closes the file if it's open.
/// </summary>
/// <returns>New reference count</returns>
size_t WofstreamSync_t::Release()
{
	EnterCriticalSection(&m_critsec);
	if (0 == --m_refCount)
	{
		// Exception handling to ensure the Leave API is called.
		try
		{
			if (m_fstream.is_open())
			{
				m_fstream.close();
			}
		}
		catch(...)
		{ }
	}
	LeaveCriticalSection(&m_critsec);
	return m_refCount;
}

// ------------------------------------------------------------------------------------------
// WofstreamManager_t
//

// Constructor (note that there should be at most one instance of this object).
WofstreamManager_t::WofstreamManager_t()
{
	InitializeCriticalSection(&m_critsec);
}

// Destructor
WofstreamManager_t::~WofstreamManager_t()
{
	DeleteCriticalSection(&m_critsec);
}

/// <summary>
/// Get a pointer to a (possibly-shared) WofstreamSync_t instance.
/// </summary>
/// <param name="szFilename">Input: requested file name</param>
/// <param name="ppWofstreamSync">Output: pointer to a (possibly-shared) WofstreamSync_t instance, if successful; nullptr if not successful.</param>
/// <param name="bAppend">Input: overwrite the file or append to one if it already exists.</param>
/// <param name="uSizeThreshold">Input: log file size threshold (0 for no maximum)</param>
/// <returns>true if successful, false otherwise</returns>
bool WofstreamManager_t::GetWofstream(const wchar_t* szFilename, WofstreamSync_t** ppWofstreamSync, bool bAppend, uint64_t uSizeThreshold)
{
	// Initialize return value and output parameter
	bool retval = false;
	*ppWofstreamSync = nullptr;
	// Serialize access to the collection
	EnterCriticalSection(&m_critsec);
	// Exception handling to ensure the Leave API gets called
	try
	{
		// Get the canonicalized name for the input file name.
		// Fail if we can't get a canonicalized name for some reason.
		std::wstring sCanonicalizedName, sCanonicalizedNameCasePreserved;
		if (CanonicalizedNames(szFilename, sCanonicalizedName, sCanonicalizedNameCasePreserved))
		{
			// Do we already have an instance corresponding to this canonicalized name?
			WofstreamSyncMap_t::iterator iter = m_wofstreams.find(sCanonicalizedName);
			if (m_wofstreams.end() == iter)
			{
				// If not, create a new instance and prepare to add it to the collection
				WofstreamSync_t* pWofstreamSync = new WofstreamSync_t;
				// Create a new std::wofstream, optionally appending vs. overwriting
				if (CreateFileOutput(szFilename, pWofstreamSync->m_fstream, bAppend))
				{
					// If successful, add the new WofstreamSync_t to the collection.
					pWofstreamSync->m_sCanonicalizedName = sCanonicalizedName;
					pWofstreamSync->m_sCanonicalizedNameCasePreserved = sCanonicalizedNameCasePreserved;
					pWofstreamSync->m_uSizeThreshold = uSizeThreshold;
					// Increment its reference count
					pWofstreamSync->AddRef();
					// Add it to the collection
					m_wofstreams[sCanonicalizedName] = pWofstreamSync;
					// Return pointer to the new object through the output parameter
					*ppWofstreamSync = pWofstreamSync;
					// Success!
					retval = true;
				}
				else
				{
					// Couldn't create the new std::wofstream; delete the newly-created object.
					delete pWofstreamSync;
				}
			}
			else
			{
				// There's already an instance for this file. Get a pointer to it.
				WofstreamSync_t* pWofstreamSync = m_wofstreams[sCanonicalizedName];
				// Increment its reference count
				pWofstreamSync->AddRef();
				// Return pointer to the existing object through the output parameter
				*ppWofstreamSync = pWofstreamSync;
				// Success!
				retval = true;
			}
		}
	}
	catch (...) {}
	LeaveCriticalSection(&m_critsec);
	return retval;
}

/// <summary>
/// Release previously returned WofstreamSync_t instance.
/// Caller should not attempt to access the pointer after passing it to this method.
/// </summary>
/// <param name="pWofstreamSync">Input: previously returned pointer to WofstreamSync_t instance.</param>
void WofstreamManager_t::ReleaseWofstream(WofstreamSync_t* pWofstreamSync)
{
	// Serialize access to the collection
	EnterCriticalSection(&m_critsec);
	// Exception handling to ensure the Leave API gets called
	try
	{
		// Decrement the reference count (closing the file if it's the last reference).
		if (0 == pWofstreamSync->Release())
		{
			// If that was the last reference, remove it from the collection...
			WofstreamSyncMap_t::iterator iter = m_wofstreams.find(pWofstreamSync->m_sCanonicalizedName);
			if (m_wofstreams.end() != iter)
			{
				m_wofstreams.erase(iter);
			}
			// ... and delete it
			delete pWofstreamSync;
		}
	}
	catch(...) {}
	LeaveCriticalSection(&m_critsec);
}

/// <summary>
/// Canonicalizes input file name so that relative vs. absolute paths, short vs. long filenames, upper vs. lower case
/// referencing the same file get the same canonicalized file name.
/// Note: implementation does not chase down overlapping drive letter mappings (e.g., NET USE, SUBST).
/// </summary>
/// <param name="szFilename">Input: file name</param>
/// <param name="sCanonicalizedName">Output: canonicalized file name</param>
/// <param name="sCanonicalizedNameCasePreserved">Output: canonicalized file name, case preserved</param>
/// <returns>true if successful; false otherwise</returns>
bool WofstreamManager_t::CanonicalizedNames(const wchar_t* szFilename, std::wstring& sCanonicalizedName, std::wstring& sCanonicalizedNameCasePreserved)
{
	sCanonicalizedName.clear();
	sCanonicalizedNameCasePreserved.clear();
	const DWORD bufsize = MAX_PATH * 2;
	wchar_t szFullPathName[bufsize] = { 0 }, szLongPathName[bufsize] = { 0 };

	// Convert a relative path to an absolute path. Note that file might or might not exist.
	DWORD ret = GetFullPathNameW(szFilename, bufsize, szFullPathName, nullptr);
	// 0 is failure; ret >= bufsize means the buffer is too small. It *is* possible to create files
	// with paths longer than MAX_PATH*2, but we're not wasting time on that here. Just fail that.
	if (0 < ret && ret < bufsize)
	{
		// If the file exists, get the long path name.
		ret = GetLongPathNameW(szFullPathName, szLongPathName, bufsize);
		// 0 is failure; ret >= bufsize means the buffer is too small. It *is* possible to create files
		// with paths longer than MAX_PATH*2, but we're not wasting time on that here. Just fail that.
		if (0 < ret && ret < bufsize)
		{
			// Got the long path; use that
			sCanonicalizedName = sCanonicalizedNameCasePreserved = szLongPathName;
		}
		else
		{
			// No long path - possibly doesn't exist. Use the already-obtained full path
			sCanonicalizedName = sCanonicalizedNameCasePreserved = szFullPathName;
		}
		// Upper-case the canonicalized name so that all future comparisons are case-insensitive.
		WString_To_Upper(sCanonicalizedName);
		// Success!
		return true;
	}
	return false;
}

