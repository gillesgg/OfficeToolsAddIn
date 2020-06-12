#include "pch.h"
#include "OfficeAddinInformation.h"
#include "ExcelProcessInformation.h"
#include "Logger.h"

#define MAX_NUM 2048

int case_insensitive_match(std::wstring s1, std::wstring s2)
{
	transform(s1.begin(), s1.end(), s1.begin(), ::tolower);
	transform(s2.begin(), s2.end(), s2.begin(), ::tolower);
	if (s1.compare(s2) == 0)
		return 1; //The strings are same
	return 0; //not matched
}

int ExcelProcessInformation::ShowModules(DWORD processID, ProcessInformation& processinformation)
{
	LOG_DEBUG << __FUNCTION__;
	
	HMODULE h_mods[1024];
	DWORD cbNeeded = 0;

	HANDLE h_process = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processID);

	if (nullptr == h_process)
	{
		LOG_DEBUG << __FUNCTION__ << " OpenProcess Failed GetLastError=" << GetLastError();
		return HRESULT_FROM_WIN32(GetLastError());
	}	
	if (EnumProcessModules(h_process, h_mods, sizeof(h_mods), &cbNeeded))
	{
		for (unsigned int i = 0; i < (cbNeeded / sizeof(HMODULE)); i++)
		{
			TCHAR szModName[MAX_PATH];
			if (GetModuleFileNameEx(h_process, h_mods[i], szModName, sizeof(szModName) / sizeof(TCHAR)))
			{
				if (szModName != nullptr)
				{
					processinformation.modules_.push_back(szModName);
					LOG_DEBUG << "module=" << szModName;
				}               
			}
		}
	}
	CloseHandle(h_process);
	return S_OK;
}

HRESULT ExcelProcessInformation::ShowProcesses(const std::wstring filter, ProcessInformation& processinformation)
{
	LOG_DEBUG << __FUNCTION__;
	
	DWORD aProcesses[MAX_NUM];
	DWORD cb_needed;

	if (!EnumProcesses(aProcesses, sizeof(aProcesses), &cb_needed))
	{    	
		LOG_DEBUG << __FUNCTION__ << " EnumProcesses Failed GetLastError=" << GetLastError();
		HRESULT_FROM_WIN32(GetLastError());
	}
	DWORD c_processes = cb_needed / sizeof(DWORD);	
	for (unsigned int i = 0; i < c_processes; i++)
	{
		if (aProcesses[i] != 0)
		{
			HANDLE h_process = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, aProcesses[i]);
			if (h_process != nullptr)
			{
				wchar_t name[2048] = { '\0' };
				GetModuleBaseName(h_process, NULL, name, sizeof(name) - 1);
				fs::path file_path = name;
				auto file = file_path.filename();

				if (case_insensitive_match(file, filter))
				{
					LOG_TRACE << "Match process=" << file.c_str();
					ShowModules(aProcesses[i], processinformation);
					processinformation.Name_ = file;
					BOOL wow;
					IsWow64Process(h_process, &wow);
					processinformation.imagetype_ = (wow == FALSE) ? ImageType::x64 : ImageType::x86;
					break;
				}
				CloseHandle(h_process);
			}
			//else
			//{
			//	//LOG_TRACE << __FUNCTION__ << " OpenProcesses Failed GetLastError=" << GetLastError();
			//}
		}
	}
	return S_OK;
}