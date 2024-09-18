// Debug output stream that can write to any or all of Windows debug stream, std::wcout, std::wcerr,
// a log file (a named std::wofstream), and message boxes on the desktops of all active users.
// Each std::endl or std::flush triggers a write to the enabled destination(s).

/*
Example usage:

	// If no other changes made, this outputs to the Windows debug stream
	dbgOut << L"Hello world." << std::endl << L"The address of wmain happens to be " << wmain << std::endl;

	// Write the text to multiple destinations
	dbgOut.WriteToWCerr(true);
	dbgOut.WriteToWCout(true);
	if (!dbgOut.WriteToFile(L".\\LogFile.txt"))
		std::wcerr << L"Failed to open log file" << std::endl;
	dbgOut << L"THIS GETS WRITTEN TO STDOUT, STDERR, AND A LOG FILE." << std::endl;

	// Create additional instances that each target different destinations.
	DbgOut_t msgboxesOnly, dbgToWCerr;
	msgboxesOnly.WriteToDebugStream(false);
	msgboxesOnly.WriteToWtsMsgBox(true);
	dbgToWCerr.WriteToDebugStream(false);
	dbgToWCerr.WriteToWCerr(true);
	msgboxesOnly << "This just goes to message boxes" << std::flush;
	dbgToWCerr << L"This just goes to stderr" << std::endl;

THREADSAFE USAGE:

	Call the .locked() function on the object before inserting data into the stream. No other threads
	that also use .locked() will be able to insert data into the stream until after the first thread
	releases access by inserting std::endl or std::flush. Example:

	dbgOut.locked() << L"At this point, we saw " << pvAddr << L" which is associated with " << szName << std::endl;

	This implementation supports having multiple instances of DbgOut_t writing to the same file.

	Note that if any code between the .locked() call and the subsequent std::flush or std::endl raises 
	an exception that doesn't crash the process, it can lock up all other threads that try to output 
	to the same object until the thread sends more data to the object and eventually inserts a std::flush
	or std::endl. For example, something like this:

		__try {
			dbgOut.locked() << L"Here's an access violation opportunity: " << *pSomething << std::endl;
		}
		__except(EXCEPTION_EXECUTE_HANDLER) {
		}

	Without the __try block, an access violation would likely crash the whole process.
	Just be careful not to insert anything into the stream that is invalid and everything will be jake.

PREPEND TIMESTAMPS:
	Call .PrependTimestamp(true) to prepend a timestamp before each line of output. Timestamp is of the format:
		yyyy-MM-dd HH:mm:ss.fff
		e.g.,
		2022-11-02 12:37:02.373
	Call .PrependTimestamp(false) to turn off prepending of timestamps.
*/


#pragma once

#include <Windows.h>
#include <iostream>
#include <sstream>
#include <fstream>
#include "WofstreamManager.h"

// ------------------------------------------------------------------------------------------
/// <summary>
/// Internal class used by DbgOut_t that implements wostream redirection to zero or more destinations.
/// </summary>
class DbgOut_InternalBufferImpl : public std::wstringbuf
{
public:
	DbgOut_InternalBufferImpl();
	~DbgOut_InternalBufferImpl();

	// Set true to turn on the destination, false to disable.
	// By default, only debug stream is enabled.
	void WriteToDebugStream(bool bWriteToDebugStream);
	void WriteToWCout(bool bWriteToWCout);
	void WriteToWCerr(bool bWriteToWCerr);
	void WriteToWtsMsgBox(bool bWriteToWtsMsgBox);
	// Give a valid file path to create a new file and begin logging to it;
	// Give a null pointer or an empty string to stop file logging.
	// Optionally append rather than overwrite.
	// Optional size threshold.
	bool WriteToFile(const wchar_t* szFilename, bool bAppend = false, uint64_t uSizeThreshold = 0);

	void WriteToHANDLE(HANDLE handle);

	// Prepend timestamp to output lines
	void PrependTimestamp(bool bPrependTimestamp);

	// Acquire/release exclusive access to the object's critical section for output.
	// A single call to the "releaseOutput" function fully releases access to the critical
	// section, even if acquireOutput had been called multiple times. This helps minimize
	// the risk of deadlock (at the potential cost of thread unsafety).
	void acquireOutput();
	void releaseOutput();

private:
	// Two critical sections to serialize access:
	// m_critsecConfig serializes configuration changes to this instance;
	// m_critsecOutput serializes access to writing the output buffer.
	// m_bAcquiredOutput is used to indicate whether output has been locked one or more times.
	CRITICAL_SECTION m_critsecConfig, m_critsecOutput;
	bool m_bAcquiredOutput = false;
	bool m_bPrependTimestamp = false;

private:
	// Override of wstringbuf member function that triggers write to destination(s)
	int sync() override;
	// Internal: release output file if writing to one
	void releaseFile();

private:
	// State - which targets are currently active
	bool m_bWriteToDebugStream, m_bWriteToWCout, m_bWriteToWCerr, m_bWriteToWtsMsgBox, m_bWriteToFile;
	// Synchronized-access file stream if writing to a file
	WofstreamSync_t* m_pStreamSync = nullptr;
	HANDLE m_handle = nullptr;

	// Single instance of a managed collection of shareable std::wofstream instances.
	// Not supportable for two wofstreams to write to the same file at the same time.
	// Heap-allocated when first class instance is created, never destroyed. Reason is that
	// we can't control the ctor/dtor order of global instances of DbgOut_InternalBufferImpl
	// vs. the destruction of a static instance of WofstreamManager_t.
	static WofstreamManager_t* st_pWofstreamMgr;

private:
	// Not implemented
	DbgOut_InternalBufferImpl(const DbgOut_InternalBufferImpl&) = delete;
	DbgOut_InternalBufferImpl& operator = (const DbgOut_InternalBufferImpl&) = delete;
};

/// <summary>
/// Another internal class used by DbgOut_t, to ensure that the wstringbuf is created before the wostream.
/// </summary>
class DbgOut_InternalBuffer
{
public:
	DbgOut_InternalBuffer() = default;
	virtual ~DbgOut_InternalBuffer() = default;
protected:
	DbgOut_InternalBufferImpl m_buf;
private:
	DbgOut_InternalBuffer(const DbgOut_InternalBuffer&) = delete;
	DbgOut_InternalBuffer& operator = (const DbgOut_InternalBuffer&) = delete;
};

// ------------------------------------------------------------------------------------------

/// <summary>
/// Instances of this std::wostream can redirect output to zero or more destinations simultaneously.
/// Writes to Windows debug stream by default.
/// </summary>
class DbgOut_t : private DbgOut_InternalBuffer, public std::wostream
{
public:
	DbgOut_t() : std::wostream(&m_buf) {}
	~DbgOut_t() = default;

	/// <summary>
	/// Calling this method acquires exclusive access to the object's output buffer until the caller
	/// inserts std::endl or std::flush into it. It returns a reference to the current object so that
	/// it can be used directly in the stream; e.g.,
	///     dbgOut.locked() << L"Information: " << pvAddr << std::endl;
	/// (Note that "exclusive access" works only if callers in other threads also use the .locked() accessor.)
	/// </summary>
	DbgOut_t& locked();

	// Set true to turn on the destination, false to disable
	// By default, only debug stream is enabled.
	void WriteToDebugStream(bool bWriteToDebugStream);
	void WriteToWCout(bool bWriteToWCout);
	void WriteToWCerr(bool bWriteToWCerr);
	void WriteToWtsMsgBox(bool bWriteToWtsMsgBox);
	// Give a valid file path to create a new file and begin logging to it;
	// Give a null pointer or an empty string to stop file logging.
	// Optionally append rather than overwrite.
	// Optional size threshold.
	bool WriteToFile(const wchar_t* szFilename, bool bAppend = false, uint64_t uSizeThreshold = 0);

	void WriteToHANDLE(HANDLE handle);

	// Prepend timestamp to output lines
	void PrependTimestamp(bool bPrependTimestamp);

private:
	DbgOut_t(const DbgOut_t&) = delete;
	DbgOut_t& operator = (const DbgOut_t&) = delete;
};



/// <summary>
/// Global instance declared. You can create other instances each with their own targets enabled. 
/// </summary>
extern DbgOut_t dbgOut;