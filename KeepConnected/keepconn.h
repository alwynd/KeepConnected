#pragma once

/**
	Author: alwyn.j.dippenaar@gmail.com
	Responsible for keeping forti client connected.
**/

#ifndef _KEEPCONN_H_
#define _KEEPCONN_H_

#include <Windows.h>
#include <strsafe.h>

#include <iostream>
#include <algorithm>
#include <stdio.h>
#include <fstream>

#include <ctime>
#include <time.h>

#include <uiautomation.h>

#include <vector>

#include <mutex>
#include <process.h>    /* _beginthread, _endthread */
#include <Psapi.h>
#pragma comment(lib, "Kernel32.lib")
#pragma comment(lib, "Psapi.lib")

using namespace std;

// A running process.
struct RunningProcess
{
	RunningProcess()
	{
		WindowHandle = NULL;
		ZeroMemory(WindowTitle, sizeof(char) * 1024);
		ZeroMemory(ImageName, sizeof(char) * 1024);
		IsMinimized = false;
	};

	bool operator < (const RunningProcess& a) const
	{
		std::string aWindowTitle(WindowTitle);
		std::string bWindowTitle(a.WindowTitle);

		return (aWindowTitle < bWindowTitle);
	}


	HWND WindowHandle;
	char WindowTitle[1024];
	char ImageName[1024];
	bool IsMinimized;
};


const long long CurrentTimeMS();							// Get current ms
void FormatDate(char* buffer);								// Format date
void debug(const char* msg);								// Debug string to console
void trace(const char* msg);								// Debug string to console
void log(const char* msg);									// log string to console
bool FileExists(const char* file);							// Does file exist.
bool Find(const char* WindowText, RunningProcess& process);
int clamp(const int& value, const int& min, const int& max); // clamps

void ProcessThread(void * parm);
void LogLastError();
void ReleaseArray(IUIAutomationElementArray *elements);
IUIAutomationElement* FindElement(IUIAutomationElementArray *controls, const char* type, const char* name);
IUIAutomationElement* FindElement(IUIAutomationElementArray *controls, const char* type, const char* name, const BOOL isEnabled);
void ListDescendants(IUIAutomation *ppAutomation, IUIAutomationElement* pParent, int indent);
void ShowElement(IUIAutomationElement* element, int index=0);
void DisconnectFortiClient(IUIAutomationElementArray *windows, IUIAutomationElement *root, IUIAutomation *ppAutomation);

std::string utf8_encode(const std::wstring &wstr);
std::wstring utf8_decode(const std::string &str);

#define BOOL_TEXT(x) (x ? "true" : "false")
#define SAFE_REL(x) if (x) x->Release(); x = NULL;
#define ERROR_DONE LogLastError(); goto done;
#define SAFE_REL_ARRAY(x) ReleaseArray(x); SAFE_REL(x);





#endif