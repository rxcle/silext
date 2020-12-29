/* Silext - A Silverlight installer extractor - Copyright (c) 2020 Rxcle 
*
* Usage: Silext <Silverlight_x64.exe> <target_path>
* Result: >= 0 Success
*          < 0 Failure
*/

#include <iostream>
#include <sstream>
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
#include <map>
#include <bitextractor.hpp>
#include <filesystem>

#pragma comment(lib, "msi.lib")
#pragma comment(lib, "setupapi.lib")
#pragma comment(lib, "Shlwapi.lib")

namespace fs = std::filesystem;

/*
STEPS:
1: Extract EXE using Z-7ip
	RESULT:
	- Silverlight.7z
	- silverlight.msi
2: Extract Silverlight.7z using 7-Zip
	RESULT:
	- Silverlight.msp
3: Extract streams from Silverlight.msp
	RESULT:
	- #oldTocurrent.mst
	- oldTocurrent.mst
	- PCW_CAB_Silver.cab
4: Apply "oldTocurrent.mst" transform to silverlight.msi from step 1 (in memory)
5: Read File and Directory tables
6: Create directory structure using Directory table info
7: Reconstruct directory structure and file names
8: Extract PCW_CAB_Silver.cab using CAB to names obtained from 6 and 7
	RESULT: Uncompressed files in structure
DONE
*/

const CLSID CLSID_MsiTransform = { 0xC1082, 0x0, 0x0, {0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46} };

struct DirInfo
{
	std::wstring ParentKey;
	std::wstring Name;
	DirInfo* ParentDir;
};

struct FileInfo
{
	std::wstring FileName;
	std::wstring DirectoryKey;
};

struct DbInfo
{
	std::map<std::wstring, DirInfo> Directories;
	std::map<std::wstring, FileInfo> Files;
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

void get_directories(PMSIHANDLE& hDatabase, std::map<std::wstring, DirInfo>& directories)
{
	execute_view(hDatabase,
		L"SELECT Directory, Directory_Parent, DefaultDir FROM Directory",
		[&directories](PMSIHANDLE& hRecord)
	{
		directories[get_record_string(hRecord, 1)] = {
			get_record_string(hRecord, 2),
			get_record_string(hRecord, 3),
			nullptr
		};
	});
}

void get_files(PMSIHANDLE& hDatabase, std::map<std::wstring, FileInfo>& files)
{
	execute_view(hDatabase,
		L"SELECT File, FileName, Directory_ FROM File, Component WHERE File.Component_ = Component.Component",
		[&files](PMSIHANDLE& hRecord)
	{
		files[get_record_string(hRecord, 1)] = {
			get_record_string(hRecord, 2),
			get_record_string(hRecord, 3)
		};
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

std::vector<std::wstring> split(const std::wstring& s, wchar_t seperator)
{
	std::vector<std::wstring> output;
	std::wstring::size_type prev_pos = 0, pos = 0;
	while ((pos = s.find(seperator, pos)) != std::wstring::npos)
	{
		auto substring(s.substr(prev_pos, pos - prev_pos));
		output.push_back(substring);
		prev_pos = ++pos;
	}

	output.push_back(s.substr(prev_pos, pos - prev_pos));
	return output;
}

void get_directory_parts(const std::map<std::wstring, DirInfo>& directories, const std::wstring& key, std::vector<std::wstring>& parts_out)
{
	auto dirInfoIt = directories.find(key);
	if (dirInfoIt == directories.end()) 
	{
		//handle the error
	}
	else 
	{
		auto& dirInfo = dirInfoIt->second;
		auto dirNameParts = split(dirInfo.Name, '|');
		auto targetDirName = dirNameParts.back().c_str();
		if (!dirInfo.ParentKey.empty())
			get_directory_parts(directories, dirInfo.ParentKey, parts_out);
		parts_out.push_back(targetDirName);
	}
}

std::wstring combine_directory_parts(const std::vector<std::wstring>& parts)
{
	std::wstringstream ss;
	std::for_each(parts.begin(), parts.end(), [&ss](const std::wstring& s) 
	{ 
		ss << s << L"\\";
	});
	return ss.str();
}

bool extract_cab(const std::wstring& cabName, const std::wstring& targetPath, const DbInfo& dbInfo)
{
	auto context = std::pair<std::wstring, DbInfo>(targetPath, dbInfo);
	return SetupIterateCabinet(cabName.c_str(), 0,
		[](PVOID context, UINT notification, UINT_PTR param1, UINT_PTR param2) -> UINT
	{
		if (notification == SPFILENOTIFY_FILEINCABINET)
		{
			auto fileInCabinetInfo = (FILE_IN_CABINET_INFO*)param1;

			auto ccontext = static_cast<std::pair<std::wstring, DbInfo>*>(context);
			DbInfo& dbInfo = ccontext->second;

			auto fileInfoIt = dbInfo.Files.find(fileInCabinetInfo->NameInCabinet);
			FileInfo& fileInfo = fileInfoIt->second;
			auto fileNameParts = split(fileInfo.FileName, '|');
			auto targetFileName = fileNameParts.back().c_str();

			std::vector<std::wstring> dirParts;
			dirParts.push_back(ccontext->first);
			get_directory_parts(dbInfo.Directories, fileInfo.DirectoryKey, dirParts);
			auto dirPath = combine_directory_parts(dirParts);

			std::error_code errorCode;
			std::filesystem::create_directories(dirPath, errorCode);
			if (errorCode)
				return FILEOP_ABORT;

			PathCombine(fileInCabinetInfo->FullTargetName, dirPath.c_str(), targetFileName);
			return FILEOP_DOIT;
		}
		return NO_ERROR;
	}, static_cast<void*>(&context));
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

bool cleanup_workdir(const std::wstring& workdir)
{
	auto tempFiles = find_files(workdir, L"*");
	std::error_code errorCode;
	for (auto& tempFile : tempFiles)
	{
		fs::permissions(tempFile,
			fs::perms::owner_write,
			fs::perm_options::add, errorCode);
		fs::remove(tempFile, errorCode);
	}
	if (!errorCode)
		fs::remove(workdir, errorCode);
	return !errorCode;
}

enum class ReturnCode
{
	Success = 0,
	SuccessNoCleanup = 1,

	CannotInitializeWorkDir = -1,
	InvalidArguments = -2,
	UnexpectedAmountOfMsiFiles = -3,
	UnexpectedAmountOf7zFiles = -4,
	UnexpectedAmountOfMspFiles = -5,
	UnexpectedAmountOfMstFiles = -6,
	UnexpectedAmountOfPayloadFiles = -7,
	UnexpectedAmountOfCabFiles = -8,
	ErrorExtractingCab = -9
};

ReturnCode extract_setup(const std::wstring& setupExeName, const std::wstring& targetPath, const std::wstring& workDir)
{
	bit7z::Bit7zLibrary blib;
	bit7z::BitExtractor cextractor(blib, bit7z::BitFormat::Cab);
	cextractor.extract(setupExeName, workDir);

	auto msiFiles = find_files(workDir, L"*.msi");
	if (msiFiles.size() != 1)
		return ReturnCode::UnexpectedAmountOfMsiFiles;

	auto sevenZipFiles = find_files(workDir, L"*.7z");
	if (sevenZipFiles.size() != 1)
		return ReturnCode::UnexpectedAmountOf7zFiles;

	bit7z::BitExtractor extractor(blib, bit7z::BitFormat::SevenZip);
	extractor.extract(sevenZipFiles.front(), workDir);

	auto mspFiles = find_files(workDir, L"*.msp");
	if (mspFiles.size() != 1)
		return ReturnCode::UnexpectedAmountOfMspFiles;

	extract_msp(mspFiles.front(), workDir);

	auto cabFiles = find_files(workDir, L"*.cab");
	if (cabFiles.size() != 1)
		return ReturnCode::UnexpectedAmountOfCabFiles;

	auto mstFiles = find_files(workDir, L"oldToCurrent.mst");
	if (mstFiles.size() != 1)
		return ReturnCode::UnexpectedAmountOfMstFiles;

	DbInfo dbInfo;
	get_files_from_mst(msiFiles.front(), mstFiles.front(), dbInfo);
	if (dbInfo.Files.empty() || dbInfo.Directories.empty())
		return ReturnCode::UnexpectedAmountOfPayloadFiles;

	if (!extract_cab(cabFiles.front(), targetPath, dbInfo))
		return ReturnCode::ErrorExtractingCab;

	return ReturnCode::Success;
}

int wmain(int argc, wchar_t* argv[])
{
	if (argc != 3)
		return static_cast<int>(ReturnCode::InvalidArguments);

	const std::wstring setupExeName = argv[1];
	const std::wstring targetPath = argv[2];

	std::error_code errorCode;
	const std::wstring workDir = concat_path(fs::temp_directory_path(errorCode), L"rxcle-silext");
	if (!errorCode)
		fs::create_directories(workDir, errorCode);
	if (errorCode)
		return static_cast<int>(ReturnCode::CannotInitializeWorkDir);

	auto extractResult = extract_setup(setupExeName, targetPath, workDir);

	bool cleanedUp = cleanup_workdir(workDir);

	return static_cast<int>(
		extractResult == ReturnCode::Success && !cleanedUp 
		? ReturnCode::SuccessNoCleanup 
		: extractResult);
}