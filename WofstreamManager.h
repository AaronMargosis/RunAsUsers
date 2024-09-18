// Manages serialization of access to possibly-shared std::wofstream instances.
//
// If two instances of std::wofstream write to the same file, the results are "undefined" at best, and
// can cause memory access violations. This code enables multiple threads to enjoy serialized access to
// possibly-shared std::wofstream instances. Filenames are canonicalized so that the code can determine
// when multiple references use different names to access the same file (e.g., long vs. short names, 
// relative vs. absolute paths, upper vs. lower case).

#pragma once

#include <Windows.h>
#include <fstream>
#include <unordered_map>

/// <summary>
/// Object that encapsulates a std::wofstream, a critical section for serializing access, the
/// canonicalized name it's referenced under, a maximum size, and a reference count.
/// </summary>
struct WofstreamSync_t
{
public:
	// public data
	std::wofstream m_fstream;
	CRITICAL_SECTION m_critsec;
	std::wstring m_sCanonicalizedName, m_sCanonicalizedNameCasePreserved;
	uint64_t m_uSizeThreshold;

public:
	// constructor, destructor
	WofstreamSync_t();
	~WofstreamSync_t();

	/// <summary>
	/// Enforces the file's size threshold (if non-zero). If the size has reached or exceeded 
	/// the threshold, closes the stream, renames the file with a timestamp in the file name, 
	/// then opens a new stream with the original file name.
	/// CALLER MUST HAVE ACQUIRED THE CRITICAL SECTION.
	/// </summary>
	void EnforceSizeThreshold();

	// For reference counting. Returns the new reference count. (If Release() returns 0, it can be deleted.)

	/// <summary>
	/// Increase this object's reference count
	/// </summary>
	/// <returns>New reference count</returns>
	size_t AddRef();

	/// <summary>
	/// Decrement this object's reference count. If decremented to 0, closes the file if it's open.
	/// </summary>
	/// <returns>New reference count</returns>
	size_t Release();

private:
	// private data
	size_t m_refCount = 0;

private:
	WofstreamSync_t(const WofstreamSync_t&) = delete;
	WofstreamSync_t& operator = (const WofstreamSync_t&) = delete;
};


/// <summary>
/// Mapping of canonicalized file name to a POINTER to a heap-allocated WofstreamSync_t.
/// </summary>
typedef std::unordered_map<std::wstring, WofstreamSync_t*> WofstreamSyncMap_t;


/// <summary>
/// Class that manages shareable instances of std::wofstream with serialized access.
/// The process should have at most one instance of this class.
/// (Could have been instantiated as a singleton standalone instance, but its lifetime must
/// exceed that of all the DbgOut_t instances that need it.)
/// </summary>
class WofstreamManager_t
{
public:
	// constructor, destructor
	WofstreamManager_t();
	~WofstreamManager_t();

	/// <summary>
	/// Get a pointer to a (possibly-shared) WofstreamSync_t instance.
	/// </summary>
	/// <param name="szFilename">Input: requested file name</param>
	/// <param name="ppWofstreamSync">Output: pointer to a (possibly-shared) WofstreamSync_t instance, if successful; nullptr if not successful.</param>
	/// <param name="bAppend">Input: overwrite the file or append to one if it already exists.</param>
	/// <param name="uSizeThreshold">Input: log file size threshold (0 for no maximum)</param>
	/// <returns>true if successful, false otherwise</returns>
	bool GetWofstream(const wchar_t* szFilename, WofstreamSync_t** ppWofstreamSync, bool bAppend = false, uint64_t uSizeThreshold = 0);

	/// <summary>
	/// Release previously returned WofstreamSync_t instance.
	/// Caller should not attempt to access the pointer after passing it to this method.
	/// </summary>
	/// <param name="pWofstreamSync">Input: previously returned pointer to WofstreamSync_t instance.</param>
	void ReleaseWofstream(WofstreamSync_t* pWofstreamSync);

private:
	/// <summary>
	/// Canonicalizes input file name so that relative vs. absolute paths, short vs. long filenames, upper vs. lower case
	/// referencing the same file get the same canonicalized file name.
	/// Note: implementation does not chase down overlapping drive letter mappings (e.g., NET USE, SUBST).
	/// </summary>
	/// <param name="szFilename">Input: file name</param>
	/// <param name="sCanonicalizedName">Output: canonicalized file name</param>
	/// <param name="sCanonicalizedNameCasePreserved">Output: canonicalized file name, case preserved</param>
	/// <returns>true if successful; false otherwise</returns>
	static bool CanonicalizedNames(const wchar_t* szFilename, std::wstring& sCanonicalizedName, std::wstring& sCanonicalizedNameCasePreserved);

private:
	// Collection that maps canonicalized file names to (possibly-shared) WofstreamSync_t instances.
	WofstreamSyncMap_t m_wofstreams;

	// Critical section to serialize access to the collection.
	CRITICAL_SECTION m_critsec;

private:
	WofstreamManager_t(const WofstreamManager_t&) = delete;
	WofstreamManager_t& operator = (const WofstreamManager_t&) = delete;
};

