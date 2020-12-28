#include <iostream>
#include <tchar.h>
#include <crtdbg.h>
#include <windows.h>
#include <objbase.h>
#include <msiquery.h>
#include <setupapi.h>
#include <fstream>
#include <vector>
#include <functional>
#include <shlwapi.h>
#include <algorithm>

#pragma comment(lib, "msi.lib")
#pragma comment(lib, "setupapi.lib")
#pragma comment(lib, "Shlwapi.lib")
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

const CLSID CLSID_MsiTransform = { 0xC1082, 0x0, 0x0, {0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46} };

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

struct DbInfo
{
	std::vector<DirInfo> Directories;
	std::vector<FileInfo> Files;
};

bool make_path(LPTSTR pszDest, size_t cchDest, const std::wstring& pszDir, const std::wstring& pszName, const std::wstring& pszExt)
{
	// Make sure pszDest is NULL-terminated.
	pszDest[0] = TEXT('\0');

	size_t len = pszDir.length();
	if (len)
	{
		if (0 != _tcsncpy_s(pszDest, cchDest, pszDir.c_str(), pszDir.length()))
		{
			return false;
		}

		if (len && TEXT('\\') != pszDest[len - 1])
		{
			// Make sure the path ends with a "\".
			if (0 != _tcsncat_s(pszDest, cchDest, TEXT("\\"), _TRUNCATE))
			{
				return false;
			}
		}
	}

	// Append the file name.
	if (0 != _tcsncat_s(pszDest, cchDest, pszName.c_str(), _TRUNCATE))
	{
		return false;
	}

	// Append the extension.
	if (!pszExt.empty())
	{
		if (0 != _tcsncat_s(pszDest, cchDest, pszExt.c_str(), _TRUNCATE))
		{
			return false;
		}
	}

	return true;
}

std::wstring get_record_string(MSIHANDLE hRecord, unsigned int iField)
{
	unsigned long cchProperty = 0;
	wchar_t szValueBuf[MAX_PATH] = L"";

	if (MsiRecordGetString(hRecord, iField, szValueBuf, &cchProperty) == ERROR_MORE_DATA)
	{
		cchProperty++;
		if (cchProperty > MAX_PATH) cchProperty = MAX_PATH;
		MsiRecordGetString(hRecord, iField, szValueBuf, &cchProperty);
	}

	return std::wstring(szValueBuf, cchProperty);
}

void execute_view(PMSIHANDLE& hDatabase, std::wstring query, std::function<void(PMSIHANDLE& hRecord)> recordFunc)
{
	PMSIHANDLE hView, hRecord;
	if (MsiDatabaseOpenView(hDatabase, query.c_str(), &hView) == ERROR_SUCCESS)
		if (MsiViewExecute(hView, NULL) == ERROR_SUCCESS)
			while (MsiViewFetch(hView, &hRecord) == ERROR_SUCCESS)
				recordFunc(hRecord);
}

void get_directories(PMSIHANDLE& hDatabase, std::vector<DirInfo>& directories)
{
	execute_view(hDatabase,
		L"SELECT Directory, Directory_Parent, DefaultDir FROM Directory",
		[&directories](PMSIHANDLE& hRecord)
	{
		directories.push_back({
			get_record_string(hRecord, 1),
			get_record_string(hRecord, 2),
			get_record_string(hRecord, 3)
			});
	});
}

void get_files(PMSIHANDLE& hDatabase, std::vector<FileInfo>& files)
{
	execute_view(hDatabase,
		L"SELECT File, FileName, Directory_ FROM File, Component WHERE File.Component_ = Component.Component",
		[&files](PMSIHANDLE& hRecord)
	{
		files.push_back({
			get_record_string(hRecord, 1),
			get_record_string(hRecord, 2),
			get_record_string(hRecord, 3)
			});
	});
}

unsigned int save_stream(MSIHANDLE hRecord, const std::wstring& directory)
{
	unsigned int uiError = NOERROR;
	wchar_t szPath[MAX_PATH];
	char szBuffer[256];
	unsigned long cbBuffer = sizeof(szBuffer);
	std::ofstream file;

	std::wstring streamName = get_record_string(hRecord, 1);
	if (!streamName.empty() && streamName[0] != 5)
	{
		// Create the local file with the simple CFile write-only class.
		do
		{
			uiError = MsiRecordReadStream(hRecord, 2, szBuffer, &cbBuffer);
			if (ERROR_SUCCESS == uiError)
			{
				if (!file.is_open())
				{
					if (0 == memcmp(szBuffer, "MSCF", 4))
					{
						if (make_path(szPath, MAX_PATH, directory.c_str(), streamName.c_str(), L".cab"))
						{
							// Create the local file in which data is written.
							_tprintf(TEXT("%s\n"), szPath);
							file.open(szPath, std::ios_base::binary);
						}
					}
				}

				file.write(szBuffer, cbBuffer);
			}
		} while (cbBuffer && ERROR_SUCCESS == uiError);
	}

	file.close();

	return uiError;
}

void save_streams(PMSIHANDLE& hDatabase, const std::wstring& directory)
{
	execute_view(hDatabase,
		L"SELECT Name, Data FROM _Streams",
		[&directory](PMSIHANDLE& hRecord)
	{
		save_stream(hRecord, directory);
	});
}

HRESULT save_storage(IStorage* const pRootStorage, const std::wstring& pszDir, const std::wstring& pszName, const std::wstring& pszExt)
{
	HRESULT hr = NOERROR;
	wchar_t szPath[MAX_PATH] = { L'\0' };
	IStorage* pStg = nullptr;
	IStorage* pFileStg = nullptr;

	_ASSERTE(pRootStorage);

	hr = pRootStorage->OpenStorage(
		pszName.c_str(),
		NULL,
		STGM_READ | STGM_SHARE_EXCLUSIVE,
		NULL,
		0,
		&pStg);
	if (SUCCEEDED(hr) && pStg)
	{
		if (!make_path(szPath, MAX_PATH, pszDir, pszName, pszExt))
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

	if (pFileStg) pFileStg->Release();
	if (pStg) pStg->Release();

	return hr;
}

void extract_msp(const std::wstring& mspName, const std::wstring& workDir)
{
	IStorage* pRootStorage = nullptr;

	//// Open the root storage file and extract storages first. Storages cannot
	//// be extracted using MSI APIs so we must use the compound file implementation
	//// for IStorage.
	HRESULT hr = StgOpenStorage(mspName.c_str(), NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pRootStorage);
	if (SUCCEEDED(hr) && pRootStorage)
	{
		IEnumSTATSTG* pEnum = nullptr;
		hr = pRootStorage->EnumElements(0, NULL, 0, &pEnum);
		if (SUCCEEDED(hr))
		{
			STATSTG stg = { 0 };
			while (S_OK == (hr = pEnum->Next(1, &stg, NULL)))
			{
				if (STGTY_STORAGE == stg.type && stg.clsid == CLSID_MsiTransform)
					save_storage(pRootStorage, workDir, stg.pwcsName, L".mst");
			}
			if (pEnum) pEnum->Release();
		}
	}
	if (pRootStorage) pRootStorage->Release();

	PMSIHANDLE hDatabase;
	if (MsiOpenDatabase(mspName.c_str(), (LPCTSTR)(MSIDBOPEN_READONLY + MSIDBOPEN_PATCHFILE), &hDatabase) == ERROR_SUCCESS)
	{
		save_streams(hDatabase, workDir);
	}
	MsiCloseHandle(hDatabase);
}

void get_files_from_mst(const std::wstring& msiName, const std::wstring& mstFile, DbInfo& dbInfo)
{
	PMSIHANDLE hDatabase = NULL;
	PMSIHANDLE hView = NULL;
	PMSIHANDLE hRecord = NULL;

	UINT dwError = MsiOpenDatabase(msiName.c_str(), MSIDBOPEN_READONLY, &hDatabase);
	if (ERROR_SUCCESS == dwError)
	{
		MsiDatabaseApplyTransform(hDatabase, mstFile.c_str(), 0);

		get_directories(hDatabase, dbInfo.Directories);
		get_files(hDatabase, dbInfo.Files);
	}
}

void extract_cab(const std::wstring& cabName, const std::wstring& targetPath)
{
	SetupIterateCabinet(cabName.c_str(), 0,
		[](PVOID context, UINT notification, UINT_PTR param1, UINT_PTR param2) -> UINT 
		{
			if (notification == SPFILENOTIFY_FILEINCABINET)
			{
				auto fileInCabinetInfo = (FILE_IN_CABINET_INFO*)param1;

				// TODO: Set full path based on Directories/Files extracted from MST
				auto tpath = (std::wstring*)context;
				PathCombine(fileInCabinetInfo->FullTargetName, tpath->c_str(), fileInCabinetInfo->NameInCabinet);

				return FILEOP_DOIT;
			}
			return NO_ERROR;
		},
		(void*)&targetPath);
}

std::wstring concat_path(const std::wstring& firstPath, const std::wstring& secondPath)
{
	wchar_t combinedPath[MAX_PATH];
	PathCombine(combinedPath, firstPath.c_str(), secondPath.c_str());
	return combinedPath;
}

std::vector<std::wstring> find_files(const std::wstring& basePath, const std::wstring& filter)
{
	std::vector<std::wstring> files;

	std::wstring fullPath = concat_path(basePath, filter);

	WIN32_FIND_DATA findData = { 0 };
	HANDLE hFind = FindFirstFile(fullPath.c_str(), &findData);
	if (INVALID_HANDLE_VALUE != hFind)
	{
		do
		{
			if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
				files.push_back(concat_path(basePath, findData.cFileName));
		} while (FindNextFile(hFind, &findData) == TRUE);

		FindClose(hFind);
	}
	return files;
}

enum class ReturnCode
{
	Success = 0,
	UnexpectedAmountOfMstFiles = 1,
	UnexpectedAmountOfPayloadFiles = 2,
	UnexpectedAmountOfCabFiles = 3
};

// Entry point.
int wmain(int argc, wchar_t* argv[])
{
	const std::wstring msiName = L"C:\\Temp\\sl5\\silverlight.msi";
	const std::wstring mspName = L"C:\\Temp\\sl5\\silverlight.msp";
	const std::wstring targetPath = L"C:\\Temp\\sl5\\work\\ext";
	const std::wstring workDir = L"C:\\Temp\\sl5\\work";

	// TODO: Extract silverlight.7z

	auto cabFiles = find_files(workDir, L"*.cab");
	if (cabFiles.size() != 1)
		return static_cast<int>(ReturnCode::UnexpectedAmountOfCabFiles);

	extract_msp(mspName, workDir);

	auto mstFiles = find_files(workDir, L"oldToCurrent.mst");
	if (mstFiles.size() != 1)
		return static_cast<int>(ReturnCode::UnexpectedAmountOfMstFiles);

	DbInfo dbInfo;
	get_files_from_mst(msiName, mstFiles.front(), dbInfo);
	if (dbInfo.Files.empty() || dbInfo.Directories.empty())
		return static_cast<int>(ReturnCode::UnexpectedAmountOfPayloadFiles);

	// TODO: Construct directory structure
	// TODO: Construct full paths for files

	extract_cab(cabFiles.front(), targetPath);

	return static_cast<int>(ReturnCode::Success);
}