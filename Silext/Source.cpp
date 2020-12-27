#include <iostream>
#include <tchar.h>
#include <stdio.h>
#include <stdlib.h>
#include <crtdbg.h>
#include <windows.h>
#include <objbase.h>
#include <msiquery.h>
#include <exception>
#include <fstream>
#include <vector>

#pragma comment(lib, "msi.lib")
/*
STEPS:
1: Extract EXE using Z-7ip (ALT: research via stream?)
	RESULT:
	- Silverlight.7z
	- silverlight.msi
2: Extract Silverlight.7z
	RESULT:
	- Silverlight.msp
3: Extract streams from Silverlight.msp
	RESULT:
	- #oldTocurrent.mst
	- oldTocurrent.mst
	- PCW_CAB_Silver.cab
4: Extract PCW_CAB_Silver.cab
	RESULT: Uncompressed files with mangled names
5: Apply oldTocurrent.mst transform to silverlight.msi from step 1 (in memory)
6: Read File and Directory tables
7: Create directory structure using Directory table info
8: Copy and rename files from result 4 to location with info from 6 and 7
DONE
*/


// "C:\Temp\sl5\silverlight.msp" /out "C:\Temp\sl5\out"

typedef struct _ARGS
{
	LPTSTR Path;
	LPTSTR Directory;
} ARGS, * LPARGS;

#define SAFE_RELEASE(ptr) if (ptr) { (ptr)->Release(); ptr = NULL; }
#define COLOR_ERROR FOREGROUND_RED | FOREGROUND_INTENSITY

EXTERN_C const CLSID CLSID_MsiPatch = { 0xC1086, 0x0, 0x0, {0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46} };

DWORD ParseArguments(int argc, _TCHAR* argv[], LPARGS args);
BOOL IsPatch(IStorage* pStorage);
HRESULT SaveStorage(IStorage* pRootStorage, LPCTSTR pszDir, PCWSTR pszName, LPCTSTR pszExt);
UINT SaveStream(MSIHANDLE hRecord, LPCTSTR pszDir);
void win32_error(DWORD dwError);
void error(LPCTSTR pszFormat, ...);
LPTSTR MakePathForData(LPTSTR pszDest, size_t cchDest, LPCTSTR pszDir, LPCTSTR pszName,
	LPCVOID pBuffer, size_t cbBuffer);
UINT GetString(MSIHANDLE hRecord, UINT iField, LPTSTR* ppszProperty, DWORD* pcchProperty);
std::wstring GetString(MSIHANDLE hRecord, UINT iField);

struct DirInfo
{
	std::wstring Key;
	std::wstring ParentKey;
	std::wstring Name;
};

struct FileInfo
{
	std::wstring Key;
	std::wstring FileName;
	std::wstring DirectoryKey;
};


void get_directories(PMSIHANDLE& hDatabase, std::vector<DirInfo>& directories)
{
	PMSIHANDLE hView, hRecord;
	if (MsiDatabaseOpenView(hDatabase, L"SELECT Directory, Directory_Parent, DefaultDir FROM Directory", &hView) == ERROR_SUCCESS)
		if (MsiViewExecute(hView, NULL) == ERROR_SUCCESS)
			while (MsiViewFetch(hView, &hRecord) == ERROR_SUCCESS)
				directories.push_back({
					GetString(hRecord, 1),
					GetString(hRecord, 2),
					GetString(hRecord, 3)
					});
}

void get_files(PMSIHANDLE& hDatabase, std::vector<FileInfo>& files)
{
	PMSIHANDLE hView, hRecord;
	if (MsiDatabaseOpenView(hDatabase, L"SELECT File, FileName, Directory_ FROM File, Component WHERE File.Component_ = Component.Component", &hView) == ERROR_SUCCESS)
		if (MsiViewExecute(hView, NULL) == ERROR_SUCCESS)
			while (MsiViewFetch(hView, &hRecord) == ERROR_SUCCESS)
				files.push_back({
					GetString(hRecord, 1),
					GetString(hRecord, 2),
					GetString(hRecord, 3)
					});
}

// Entry point.
int _tmain(int argc, _TCHAR* argv[])
{
	DWORD dwError = NOERROR;
	HRESULT hr = NOERROR;
	ARGS args = { 0 };
	IStorage* pRootStorage = NULL;
	IEnumSTATSTG* pEnum = NULL;
	LPCTSTR pszPersist = MSIDBOPEN_READONLY;// +MSIDBOPEN_PATCHFILE;
	STATSTG stg = { 0 };
	PMSIHANDLE hDatabase = NULL;
	PMSIHANDLE hView = NULL;
	PMSIHANDLE hRecord = NULL;

	//dwError = ParseArguments(argc, argv, &args);
	//if (ERROR_SUCCESS != dwError)
	//{
	//	return dwError;
	//}

	LPCWSTR mspName = L"C:\\Temp\\sl5\\silverlight.msp";
	LPCWSTR workDir = L"C:\\Temp\\sl5\\work";

	//// Open the root storage file and extract storages first. Storages cannot
	//// be extracted using MSI APIs so we must use the compound file implementation
	//// for IStorage.
	hr = StgOpenStorage(
		mspName,
		NULL,
		STGM_READ | STGM_SHARE_EXCLUSIVE,
		NULL,
		0,
		&pRootStorage);
	if (SUCCEEDED(hr) && pRootStorage)
	{
	//	// Determine if the file path specifies an MSP file.
	//	// This will be used later to open the database with MSI APIs.
	//	if (IsPatch(pRootStorage))
	//	{
	//		pszPersist = MSIDBOPEN_READONLY + MSIDBOPEN_PATCHFILE;
	//	}

		hr = pRootStorage->EnumElements(0, NULL, 0, &pEnum);
		if (SUCCEEDED(hr))
		{
			while (S_OK == (hr = pEnum->Next(1, &stg, NULL)))
			{
				if (STGTY_STORAGE == stg.type)
				{
					hr = SaveStorage(pRootStorage, workDir, stg.pwcsName, TEXT(".mst"));
					if (FAILED(hr))
					{
						break;
					}
				}
				else if (STGTY_STREAM == stg.type)
				{
					// TODO: Save stream
				}
			}
			SAFE_RELEASE(pEnum);
		}
	}
	SAFE_RELEASE(pRootStorage);


	// Now open the database using MSI APIs. Patches cannot be opened simultaneously
	// since exclusive access is required and no MSI APIs are exported that accept
	// an IStorage pointer.
	//if (SUCCEEDED(hr))
	//{

	LPCWSTR dbName = L"C:\\Temp\\sl5\\silverlight.msi";

	dwError = MsiOpenDatabase(dbName, pszPersist, &hDatabase);
	if (ERROR_SUCCESS == dwError)
	{
		int tresult = MsiDatabaseApplyTransform(hDatabase, L"C:\\Temp\\sl5\\oldTocurrent.mst", 0);
		//int cresult = MsiDatabaseCommit(hDatabase);

		std::vector<DirInfo> directories;

		get_directories(hDatabase, directories);

		for (DirInfo& d : directories)
		{
			std::wcout << d.Key << L"\t" << d.ParentKey << L"\t" << d.Name << std::endl;
		}

		std::vector<FileInfo> files;

		get_files(hDatabase, files);

		for (FileInfo& f : files)
		{
			std::wcout << f.Key << L"\t" << f.FileName << L"\t" << f.DirectoryKey << std::endl;
		}

		//dwError = MsiDatabaseOpenView(hDatabase,
		//	TEXT("SELECT `Name`, `Data` FROM `_Streams`"), &hView);
		//if (ERROR_SUCCESS == dwError)
		//{
		//	dwError = MsiViewExecute(hView, NULL);
		//	if (ERROR_SUCCESS == dwError)
		//	{
		//		while (ERROR_SUCCESS == (dwError = MsiViewFetch(hView, &hRecord)))
		//		{
		//			dwError = SaveStream(hRecord, args.Directory);
		//			if (ERROR_SUCCESS != dwError)
		//			{
		//				break;
		//			}
		//		}

		//		// If there are no more records indicate success.
		//		if (ERROR_NO_MORE_ITEMS == dwError)
		//		{
		//			dwError = ERROR_SUCCESS;
		//		}
		//	}
		//}
	}
	//}

	// If a Win32 error has occurred return only the win32 error portion.
	if (FACILITY_WIN32 == HRESULT_FACILITY(hr))
	{
		dwError = HRESULT_CODE(hr);
	}
	else if (FAILED(hr))
	{
		// Just set it to the HRESULT. Many common HRESULTs
		// will yield an error string from FormatMessage.
		dwError = hr;
	}

	// Print the error to the console.
	if (ERROR_SUCCESS != dwError)
	{
		win32_error(dwError);
	}

	return dwError;
}

// Parse the file path, and optionally output directory and extension guessing switch.
DWORD ParseArguments(int argc, _TCHAR* argv[], LPARGS args)
{
	_ASSERTE(argv);
	_ASSERTE(args);

	int iParamIndex = 1;

	// Validate arguments.
	if (2 > argc)
	{
		error(TEXT("Error: you must specify a Windows Installer file from which to extract files.\n"));
		return ERROR_INVALID_PARAMETER;
	}

	if (0 == _tcsicmp(TEXT("/?"), argv[iParamIndex]) || 0 == _tcsicmp(TEXT("-?"), argv[iParamIndex]))
	{
		// Display the usage text.
		return ERROR_SUCCESS;
	}
	else if (TEXT('/') == argv[iParamIndex][0] || TEXT('-') == argv[iParamIndex][0])
	{
		// Filename should not begin with a command-switch character.
		error(TEXT("Error: invalid file name.\n"));
		return ERROR_INVALID_PARAMETER;
	}
	else
	{
		// Set the path argument.
		args->Path = const_cast<LPTSTR>(argv[iParamIndex]);
	}

	// Get the output directory if requested.
	while (++iParamIndex < argc)
	{
		// The directory in which files are extracted.
		if (0 == _tcsicmp(TEXT("/out"), argv[iParamIndex]) || 0 == _tcsicmp(TEXT("-out"), argv[iParamIndex]))
		{
			if (++iParamIndex < argc && TEXT('/') != argv[iParamIndex][0] && TEXT('-') != argv[iParamIndex][0])
			{
				args->Directory = const_cast<LPTSTR>(argv[iParamIndex]);
			}
			else
			{
				error(TEXT("Error: you must specify an output directory with /out.\n"));
				return ERROR_INVALID_PARAMETER;
			}
		}
		else
		{
			error(TEXT("Error: unknown option: %s.\n"), argv[iParamIndex]);
			return ERROR_INVALID_PARAMETER;
		}
	}

	return ERROR_SUCCESS;
}

// Colors errors on the console red and prints the formatted error.
// You can use positional format specifiers with CRT8.
void error(LPCTSTR pszFormat, ...)
{
	_ASSERTE(pszFormat);

	CONSOLE_SCREEN_BUFFER_INFO csbi;
	HANDLE hStdErr = INVALID_HANDLE_VALUE;
	va_list args;

	// Set console colors to error values.
	hStdErr = GetStdHandle(STD_ERROR_HANDLE);
	if (INVALID_HANDLE_VALUE != hStdErr)
	{
		if (GetConsoleScreenBufferInfo(hStdErr, &csbi))
		{
			// Set new console colors.
			SetConsoleTextAttribute(hStdErr, COLOR_ERROR);
		}
	}

	// Print error.
	va_start(args, pszFormat);
	_vftprintf_p(stderr, pszFormat, args);
	va_end(args);

	// Reset the console colors to original values.
	if (INVALID_HANDLE_VALUE != hStdErr)
	{
		SetConsoleTextAttribute(hStdErr, csbi.wAttributes);
	}
}

// Wrapper around FormatMessage for getting error text.
// Calls error() to print the error to the console.
void win32_error(DWORD dwError)
{
	LPTSTR pszError;

	// Format the error. Error ends with new line.
	if (FormatMessage(
		FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
		NULL,
		dwError,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		(LPTSTR)&pszError,
		0,
		NULL))
	{
		error(TEXT("Error 0x%1$08x (%1$d): %2$s"), dwError, pszError);
		LocalFree(pszError);
	}
}

// Creates a patch from components, using the current working
// directory if pszDir is NULL.
// pszExt should be either NULL or start with a ".".
LPTSTR MakePath(LPTSTR pszDest, size_t cchDest, LPCTSTR pszDir, LPCTSTR pszName, LPCTSTR pszExt)
{
	size_t len = 0;

	_ASSERTE(pszDest);
	_ASSERTE(cchDest);
	_ASSERTE(pszName);

	// Make sure pszDest is NULL-terminated.
	pszDest[0] = TEXT('\0');

	if (pszDir)
	{
		// Get the length of pszDir.
		len = _tcslen(pszDir);

		if (len && 0 != _tcsncpy_s(pszDest, cchDest, pszDir, len))
		{
			return NULL;
		}

		if (len && TEXT('\\') != pszDest[len - 1])
		{
			// Make sure the path ends with a "\".
			if (0 != _tcsncat_s(pszDest, cchDest, TEXT("\\"), _TRUNCATE))
			{
				return NULL;
			}
		}
	}

	// Append the file name.
	if (0 != _tcsncat_s(pszDest, cchDest, pszName, _TRUNCATE))
	{
		return NULL;
	}

	// Append the extension.
	if (pszExt)
	{
		if (0 != _tcsncat_s(pszDest, cchDest, pszExt, _TRUNCATE))
		{
			return NULL;
		}
	}

	return pszDest;
}

// Wrapper around allocating and filling a buffer using MsiRecordGetString().
UINT GetString(MSIHANDLE hRecord, UINT iField, LPTSTR* ppszProperty, DWORD* pcchProperty)
{
	_ASSERTE(hRecord);
	_ASSERTE(iField > 0);
	_ASSERTE(ppszProperty);
	_ASSERTE(pcchProperty);

	UINT iErr = NOERROR;
	DWORD cchProperty = 0;
	WCHAR szValueBuf[1] = L"";

	iErr = MsiRecordGetString(hRecord, iField, szValueBuf, &cchProperty);
	if (ERROR_MORE_DATA == iErr)
	{
		*ppszProperty = new TCHAR[++cchProperty];
		*pcchProperty = cchProperty;

		iErr = MsiRecordGetString(hRecord, iField, *ppszProperty, &cchProperty);
		if (ERROR_SUCCESS != iErr)
		{
			delete[] * ppszProperty;
			*ppszProperty = NULL;
			*pcchProperty = 0;
		}
	}

	return iErr;
}

// Wrapper around allocating and filling a buffer using MsiRecordGetString().
std::wstring GetString(MSIHANDLE hRecord, UINT iField)
{
	const int MaxLength = 255;
	DWORD cchProperty = 0;
	WCHAR szValueBuf[MaxLength] = L"";

	if (MsiRecordGetString(hRecord, iField, szValueBuf, &cchProperty) == ERROR_MORE_DATA)
	{
		cchProperty++;
		if (cchProperty > MaxLength) cchProperty = MaxLength;
		MsiRecordGetString(hRecord, iField, szValueBuf, &cchProperty);
	}

	return std::wstring(szValueBuf, cchProperty);
}


// Determines if the given IStorage* is for a patch
// using the STATSTG for the IStorage object.
BOOL IsPatch(IStorage* pStorage)
{
	_ASSERTE(pStorage);

	HRESULT hr = NOERROR;
	STATSTG stg = { 0 };

	hr = pStorage->Stat(&stg, STATFLAG_NONAME);
	if (SUCCEEDED(hr))
	{
		return !memcmp(&stg.clsid, &CLSID_MsiPatch, sizeof(CLSID));
	}

	return FALSE;
}

// Creates a new storage file and saves the named sub-storage of pRootStorage
// to the new storage file.
HRESULT SaveStorage(IStorage* pRootStorage, LPCTSTR pszDir, PCWSTR pszName, LPCTSTR pszExt)
{
	HRESULT hr = NOERROR;
	TCHAR szPath[MAX_PATH] = { TEXT('\0') };
	IStorage* pStg = NULL;
	IStorage* pFileStg = NULL;

	_ASSERTE(pRootStorage);
	_ASSERTE(pszName);

	hr = pRootStorage->OpenStorage(
		pszName,
		NULL,
		STGM_READ | STGM_SHARE_EXCLUSIVE,
		NULL,
		0,
		&pStg);
	if (SUCCEEDED(hr) && pStg)
	{
		if (!MakePath(szPath, MAX_PATH, pszDir, pszName, pszExt))
		{
			hr = E_INVALIDARG;
		}
		else
		{
			_tprintf(TEXT("%s\n"), szPath);

			// Create the storage file.
			hr = StgCreateDocfile(
				szPath,
				STGM_WRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE,
				0,
				&pFileStg);
			if (SUCCEEDED(hr) && pFileStg)
			{
				hr = pStg->CopyTo(0, NULL, NULL, pFileStg);
			}
		}
	}

	SAFE_RELEASE(pFileStg);
	SAFE_RELEASE(pStg);

	if (FAILED(hr))
	{
		error(TEXT("Error: failed to save storage '%s'.\n"), pszName);
	}

	return hr;
}

UINT SaveStream(MSIHANDLE hRecord, LPCTSTR pszDir)
{
	UINT uiError = NOERROR;
	TCHAR szPath[MAX_PATH];
	LPTSTR pszName = NULL;
	DWORD cchName = 0;
	CHAR szBuffer[256];
	DWORD cbBuffer = sizeof(szBuffer);
	std::ofstream file;

	try
	{
		// Get the name of the stream but skip if \005SummaryInformation stream.
		if (ERROR_SUCCESS == GetString(hRecord, 1, &pszName, &cchName) &&
			0 != _tcsncmp(pszName, TEXT("\005"), 1))
		{
			// Create the local file with the simple CFile write-only class.
			do
			{
				uiError = MsiRecordReadStream(hRecord, 2, szBuffer, &cbBuffer);
				if (ERROR_SUCCESS == uiError)
				{
					if (!file.is_open())
					{
						// Create the file path if the file is not created and assume the extension
						// if requested by fIncludeExt.
						if (!MakePathForData(szPath, MAX_PATH, pszDir, pszName, szBuffer, cbBuffer))
						{
							throw std::exception("Could not create the output path name");
						}

						// Create the local file in which data is written.
						_tprintf(TEXT("%s\n"), szPath);
						file.open(szPath, std::ios_base::binary);
					}

					file.write(szBuffer, cbBuffer);
				}
				else
				{
					throw std::exception("Could not read from stream.");
				}
			} while (cbBuffer);
		}
	}
	catch (std::exception& ex)
	{
		error(TEXT("Error: %s\n"), ex.what());
		uiError = ERROR_CANNOT_MAKE;
	}

	file.close();
	if (pszName)
	{
		delete[] pszName;
		pszName = NULL;
	}

	return uiError;
}

// Creates a patch for the given file using MakePath, but uses what of the
// buffer it can to guess the file type and infer a common file extension.
LPTSTR MakePathForData(LPTSTR pszDest, size_t cchDest, LPCTSTR pszDir, LPCTSTR pszName,
	LPCVOID pBuffer, size_t cbBuffer)
{
	LPCWSTR pszExt = NULL;
	// Cabinet (*.cab) files.
	if (0 == memcmp(pBuffer, "MSCF", 4))
	{
		pszExt = TEXT(".cab");
	}

	// Executable files. Assumed to be .dll (more common).
	else if (0 == memcmp(pBuffer, "MZ", 2))
	{
		pszExt = TEXT(".dll");
	}

	// Icon (*.ico) files. Only assumed because they're common.
	else if (0 == memcmp(pBuffer, "\0\0\1\0", 4))
	{
		pszExt = TEXT(".ico");
	}

	// Bitmap (*.bmp) files.
	else if (0 == memcmp(pBuffer, "BM", 2))
	{
		pszExt = TEXT(".bmp");
	}

	// GIF (*.gif) files.
	else if (0 == memcmp(pBuffer, "GIF", 3))
	{
		pszExt = TEXT(".gif");
	}

	// PING (*.png) files.
	else if (0 == memcmp(pBuffer, "\x89PNG", 4))
	{
		pszExt = TEXT(".png");
	}

	// TIFF (*.tif) files.
	else if (0 == memcmp(pBuffer, "II", 2))
	{
		pszExt = TEXT(".tif");
	}

	return MakePath(pszDest, cchDest, pszDir, pszName, pszExt);
}
