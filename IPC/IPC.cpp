// IPC.cpp : Defines the entry point for the console application.
// https://www.codeproject.com/Articles/1842/A-newbie-s-elementary-guide-to-spawning-processes

#include "stdafx.h"
#include "Windows.h"
#include <shellapi.h>
#include <tchar.h>

//#pragma comment(lib "shell32.lib")

int main()
{
	//int result = system("C:\\Program Files\\Program.exe");
	//ShellExecute(NULL, _T("Open"), _T("C:\\Program Files\\My Prgram\\test1.pdf"),\
		NULL, NULL, SW_SHOWNORMAL);
	SHELLEXECUTEINFO ShExecInfo = { 0 };
	ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
	ShExecInfo.hwnd = NULL;
	ShExecInfo.lpVerb = NULL;
	ShExecInfo.lpFile = _T("TestBinding.exe");
	ShExecInfo.lpParameters = _T("");
	ShExecInfo.lpDirectory = NULL;
	ShExecInfo.nShow = SW_HIDE;
	ShExecInfo.hInstApp = NULL;
	ShellExecuteEx(&ShExecInfo);
	WaitForSingleObject(ShExecInfo.hProcess, INFINITE);
    return 0;
}

