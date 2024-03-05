#include "keepconn.h"

#include <locale>

/**
	Author: alwyn.j.dippenaar@gmail.com
	The impl for keepconn
**/

bool quit = false;
std::vector<RunningProcess> Processes;
std::vector<RunningProcess> ChildElements;

bool done				= false;	// is the process done 
bool verbose			= false;    // log verbose 
bool veryverbose		= false;    // log very verbose 
									   
bool fortclient			= false;    // process forticlient reconnect
bool globalprotect  	= false;    // process globalprotect reconnect
bool dumpall			= false;    // dumpall windows, childwindows and all of their discoverable screen/UI their components/elements
bool visiblewindowsonly = false;    // only perform the task on visible windows.
std::string imagematch  = "";		// only case-insensitive match images/processes with this name. ( '' defaults to all )


bool firsttime			= true;		// first time connecting?

int timer_in_ms = 15000;
std::string logprefix = "";

void LogLastError()
{
	// Retrieve the system error message for the last-error code
	LPVOID lpMsgBuf;
	LPVOID lpDisplayBuf;
	DWORD dw = GetLastError();
	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, dw, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&lpMsgBuf, 0, NULL);

	// Display the error message and exit the process
	lpDisplayBuf = (LPVOID)LocalAlloc(LMEM_ZEROINIT, (lstrlen((LPCTSTR)lpMsgBuf) + 40) * sizeof(TCHAR));
	StringCchPrintf((LPTSTR)lpDisplayBuf, LocalSize(lpDisplayBuf) / sizeof(TCHAR), TEXT("GetLastError: ERROR %d: %s"), dw, lpMsgBuf);
	debug((LPCTSTR)lpDisplayBuf);

	LocalFree(lpMsgBuf);
	LocalFree(lpDisplayBuf);
}

// Get current ms
const long long CurrentTimeMS()
{
	chrono::system_clock::time_point currentTime = chrono::system_clock::now();

	long long transformed = currentTime.time_since_epoch().count() / 1000000;
	long long millis = transformed % 1000;

	return millis;
}

// Format date
void FormatDate(char* buffer)
{
	buffer[0] = '\0';
	time_t timer;
	tm tm_info;

	timer = time(NULL);
	localtime_s(&tm_info, &timer);

	strftime(buffer, 26, "%Y-%m-%d %H:%M:%S\0", &tm_info);
}

// Debug string to console
void debug(const char* msg)
{
	if (!verbose) return;
	log(msg);
}


void trace(const char* msg)
{
	if (!veryverbose) return;
	log(msg);
}

void log(const char* msg)
{
	char buffer[27];
	ZeroMemory(buffer, 27);
	FormatDate(buffer);

	printf("%s - %s\n", buffer, msg); fflush(stdout);
	
}


// Enum window callback
BOOL CALLBACK EnumWindowCallback(HWND hWnd, LPARAM lparam)
{
	if (!lparam)
	{
		debug("MyWorkspace.EnumWindowCallback ERROR : lparam is INVALID.");
		return TRUE;
	}
	std::vector<RunningProcess>* processes = (std::vector<RunningProcess>*) lparam;

	int length = GetWindowTextLength(hWnd);
	char buffer[1024];
	ZeroMemory(buffer, 1024);
	DWORD dwProcId = 0;

	// List visible windows with a non-empty title
	if (/*IsWindowVisible(hWnd) &&*/ length != 0)
	{
		RunningProcess process;

		process.WindowHandle = hWnd;
		int wlen = length + 1;
		GetWindowText(hWnd, process.WindowTitle, wlen);
		GetWindowThreadProcessId(hWnd, &dwProcId);
		if (dwProcId)
		{
			HANDLE hProc = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, dwProcId);
			if (hProc)
			{
				HMODULE hMod;
				DWORD cbNeeded;
				process.IsMinimized = IsIconic(hWnd);
				if (EnumProcessModules(hProc, &hMod, sizeof(hMod), &cbNeeded))
				{
					GetModuleBaseNameA(hProc, hMod, process.ImageName, MAX_PATH);
				} //if
				CloseHandle(hProc);
				processes->push_back(process);		// COPY, this is NOT BY REF
				sprintf_s(buffer, "EnumWindowCallback, Window: %p: '%-55s', image: '%-55s', isMin: '%-5s'\0", process.WindowHandle, process.WindowTitle, process.ImageName, BOOL_TEXT(process.IsMinimized)); debug(buffer);
			} //if
		} //if
	}

	return TRUE;
}




// Does file exist.
bool FileExists(const char* file)
{
	FILE* f = NULL;
	fopen_s(&f, file, "rb");
	if (f)
	{
		fclose(f);
		return  true;
	} //if

	return false;
}



BOOL CALLBACK EnumChildWindowCallback(HWND hWnd, LPARAM lparam)
{

	int length = GetWindowTextLength(hWnd);
	char buffer[1024];
	ZeroMemory(buffer, 1024);

	if (length != 0)
	{
		RunningProcess process;
		process.WindowHandle = hWnd;

		int wlen = length + 1;
		GetWindowTextA(hWnd, process.WindowTitle, wlen);
		sprintf_s(buffer, "EnumChildWindowCallback: %s: hwnd: %p, WindowText: %s\0", logprefix.c_str(), hWnd, process.WindowTitle); debug(buffer);

		ChildElements.push_back(process);
	} //if


	return TRUE;
}



// Convert a wide Unicode string to an UTF8 string
std::string utf8_encode(const std::wstring &wstr)
{
	if (wstr.empty()) return std::string();
	int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
	std::string strTo(size_needed, 0);
	WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
	return strTo;
}



// Convert an UTF8 string to a wide Unicode String
std::wstring utf8_decode(const std::string &str)
{
	if (str.empty()) return std::wstring();
	int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
	std::wstring wstrTo(size_needed, 0);
	MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
	return wstrTo;
}


void ReleaseArray(IUIAutomationElementArray *elements)
{
	char buffer[1024];
	if (elements)
	{
		int cnt = 0;
		int released = 0;
		elements->get_Length(&cnt);
		sprintf_s(buffer, "Releasing IUIElements: #%i\0", cnt); debug(buffer);
		for (int i = 0; i < cnt; i++)
		{
			IUIAutomationElement *element = NULL;
			if (elements->GetElement(i, &element) == S_OK && element)
			{
				released += 1;
				SAFE_REL(element);
			} //if
		} //for
		sprintf_s(buffer, "Released IUIElements: #%i\0", released); debug(buffer);


	} //if

}


void ShowElement(IUIAutomationElement* element, const int index)
{
	char buffer[2048];
	
	BSTR bstr;
	std::string cstrname;
	std::string cstrclass;
	std::string cstrtype;
	std::string helptext;

	element->get_CurrentName(&bstr); if (SysStringLen(bstr) > 0) cstrname = utf8_encode(bstr); SysFreeString(bstr);
	element->get_CurrentClassName(&bstr); if (SysStringLen(bstr) > 0) cstrclass = utf8_encode(bstr); SysFreeString(bstr);
	BOOL isControl = FALSE;  element->get_CurrentIsControlElement(&isControl);
	BOOL isContent = FALSE;  element->get_CurrentIsContentElement(&isContent);
	UIA_HWND hwnd = NULL; element->get_CurrentNativeWindowHandle(&hwnd);
	element->get_CurrentLocalizedControlType(&bstr); if (SysStringLen(bstr) > 0) cstrtype = utf8_encode(bstr);SysFreeString(bstr);
	element->get_CurrentHelpText(&bstr); if (SysStringLen(bstr) > 0) helptext = utf8_encode(bstr);SysFreeString(bstr);
	sprintf_s(buffer, "Element #%4i: %s: hwnd: %p, Class: %-32s, Type: %-32s, Ctl?: %5s, Content?: %5s, Text: %-64s, Help: %-64s\0", index, logprefix.c_str(), (HWND)hwnd, cstrclass.c_str(), cstrtype.c_str(), BOOL_TEXT(isControl), BOOL_TEXT(isContent), cstrname.c_str(), helptext.c_str()); debug(buffer);
	
}

void ShowElements(IUIAutomationElementArray *elements)
{
	char buffer[2048];
	
	int count = 0;
	elements->get_Length(&count);
	sprintf_s(buffer, "Got Elements: #%i %s \0", count, logprefix.c_str()); trace(buffer);
	for (int i = 0; i < count; i++)
	{
		IUIAutomationElement *element = NULL;
		if (elements->GetElement(i, &element) == S_OK && element)
		{
			ShowElement(element, i);
		} //if
	} //for

}



IUIAutomationElement* FindElement(IUIAutomationElementArray *controls, const char* type, const char* name)
{
	IUIAutomationElement *control = NULL;

	int count = 0;
	controls->get_Length(&count);
	for (int i = 0; i < count; i++)
	{
		control = NULL;
		if (controls->GetElement(i, &control) == S_OK && control)
		{
			BSTR bstr;
			std::string cstrname;
			std::string cstrtype;
			control->get_CurrentName(&bstr); if (SysStringLen(bstr) > 0) cstrname = utf8_encode(bstr); SysFreeString(bstr);
			control->get_CurrentLocalizedControlType(&bstr); if (SysStringLen(bstr) > 0) cstrtype = utf8_encode(bstr); SysFreeString(bstr);

			if (strcmp(type, cstrtype.c_str()) == 0 && strcmp(name, cstrname.c_str()) == 0)
			{
				return control;
			} //if
		} //if
	} //for

	// not found.
	return NULL;

}

IUIAutomationElement* FindElement(IUIAutomationElementArray *controls, const char* type, const char* name, const BOOL isEnabled)
{
	IUIAutomationElement *control = FindElement(controls, type, name);
	if (control)
	{
		BOOL enabled = false;
		control->get_CurrentIsEnabled(&enabled);
		if (enabled == isEnabled)
		{
			return control;
		} //if
	} //if
	return NULL;
}



void ProcessFortiClient(IUIAutomationElementArray *windows, const HWND hWndTarget, IUIAutomation *ppAutomation)
{

	IUIAutomationElement *window = NULL;			// the window to automate
	IUIAutomationElementArray *controls = NULL;		// window controls
	IUIAutomationElement *button = NULL;			// button
	IUIAutomationCondition *condition = NULL;		// always true

	if (!hWndTarget || !ppAutomation) return;
	debug("ProcessFortiClient:-- START");
	ChildElements.clear();
	logprefix = "FortiClient";
	EnumChildWindows(hWndTarget, EnumChildWindowCallback, NULL);	// win32 OLD, not used/required.

	if (ppAutomation->CreateTrueCondition(&condition) != S_OK || !condition)
	{
		debug("ProcessFortiClient unable to create TRUE condition (CreateTrueCondition) FAILED."); ERROR_DONE
	} //if

	{
		int count = 0;
		windows->get_Length(&count);
		for (int i = 0; i < count; i++)
		{
			IUIAutomationElement *element = NULL;
			if (windows->GetElement(i, &element) == S_OK && element)
			{
				UIA_HWND hwnd = NULL; element->get_CurrentNativeWindowHandle(&hwnd);
				if (hWndTarget == hwnd)
				{
					window = element;

					if (window)
					{
						int count = 0;
						if (firsttime)
						{
							firsttime = false;
							window->SetFocus();
						} //if
						// list all controls.
						if (window->FindAll(TreeScope::TreeScope_Descendants, condition, &controls) != S_OK || !controls)
						{
							debug("ProcessFortiClient FindAll window TreeScope_Descendants FAILED."); ERROR_DONE
						} //if

						debug("\nProcessFortiClient Window Controls");
						controls->get_Length(&count);
						ShowElements(controls);

						// Find the Button
						button = FindElement(controls, "button", "Connect", TRUE);
						if (button)
						{
							debug("ProcessFortiClient Found Connect button...OK");				// we are not connected. check is active/clickable.
							IUIAutomationInvokePattern* iinvoke = NULL;
							if (button->GetCurrentPattern(UIA_InvokePatternId, reinterpret_cast<IUnknown**>(&iinvoke)) == S_OK && iinvoke)
							{
								iinvoke->Invoke();
							} //if
						} //if
						else
						{
							debug("ProcessFortiClient Connect button...Not found..");			// ignore and done for now.
						} //else
					} //if
					SAFE_REL(window);
					
				} //if
				SAFE_REL(element);
			} //if
		} //for
	} //scope


done:
	SAFE_REL(condition);
	debug("ProcessFortiClient Releasing window controls."); SAFE_REL_ARRAY(controls);
	
}




bool Find(const char* WindowText, RunningProcess& process)
{
	for (int i = 0; i < Processes.size(); i++)
	{
		process = Processes[i];
		if (strcmp(process.WindowTitle, WindowText) == 0)
		{
			return true;
		} //if
	} //for
	return false;
}


int clamp(const int& value, const int& min, const int& max)
{
	return (value < min ? min : (value > max ? max : value));
}


void DumpAll(IUIAutomationElementArray *windows, IUIAutomationElement *root, IUIAutomation *ppAutomation)
{
	if (!dumpall) return;
	debug("DumpAll:-- START");

	char buffer[1024];
	IUIAutomationElement *window = NULL;			// the window to automate
	IUIAutomationElement *element = NULL;			// the window to automate
	IUIAutomationElementArray *controls = NULL;		// window controls
	IUIAutomationElementArray *childwindows = NULL;		// window controls
	
	IUIAutomationCondition* namefind = NULL;
	IUIAutomationCondition *condition = NULL;		// always true

	// a true condition to return all elements using UIAutomation API
	if (ppAutomation->CreateTrueCondition(&condition) != S_OK || !condition)
	{
		debug("DumpAll unable to create TRUE condition (CreateTrueCondition) FAILED."); ERROR_DONE
	} //if

	// iterate all running processes with windows
	for (int i = 0; i < Processes.size(); i++)
	{
		RunningProcess& process = Processes[i];
		if (process.WindowHandle)	// non null window handle?
		{
			// visible?
			if (visiblewindowsonly)
			{
				//if (!IsWindowVisible(process.WindowHandle)) continue; // next window
				if (IsWindowVisible(process.WindowHandle))
				{
					sprintf_s(buffer, "DumpAll: WINDOW IS VISIBLE: '%-55s', %p: '%s'\0", process.ImageName, process.WindowHandle, process.WindowTitle); trace(buffer);
				} //if
				else
				{
					sprintf_s(buffer, "DumpAll: WINDOW NOT VISIBLE '%-55s', %p: '%s'\0", process.ImageName, process.WindowHandle, process.WindowTitle); trace(buffer);
					continue; //for
				} //else
			} //if

			//image match?
			if (imagematch.length() > 0)
			{
				std::string lookfor(imagematch.c_str()); 
				std::transform(lookfor.begin(), lookfor.end(), lookfor.begin(), [](const unsigned char c){ return std::tolower(c, std::locale()); });

				std::string lookthrough(process.ImageName);
				std::transform(lookthrough.begin(), lookthrough.end(), lookthrough.begin(), [](const unsigned char c){ return std::tolower(c, std::locale()); });

				if (lookthrough.find(lookfor) == std::string::npos)
				{
					sprintf_s(buffer, "DumpAll: NO IMAGE MATCH '%-55s', %p: '%s'\0", process.ImageName, process.WindowHandle, process.WindowTitle); trace(buffer);
					continue; //for
				} //if
			} //if

			// @TODO temp
			//continue;
			
			// find the window
			VARIANT namevalue;
			VariantInit(&namevalue);
			namevalue.vt = VT_INT;
			namevalue.intVal = (int)process.WindowHandle;
			sprintf_s(buffer, "DumpAll: looking for img: '%-55s', window: %p: '%s'\0", process.ImageName, process.WindowHandle, process.WindowTitle); trace(buffer);
			if (ppAutomation->CreatePropertyCondition(UIA_NativeWindowHandlePropertyId, namevalue, &namefind) == S_OK && namefind)
			{
				if (root->FindAll(TreeScope::TreeScope_Descendants, namefind, &childwindows) == S_OK && childwindows)
				{
					sprintf_s(buffer, "DumpAll: UIAUTOMATION WINDOWS img: '%-55s', window: %p: '%s'\0", process.ImageName, process.WindowHandle, process.WindowTitle);
					logprefix.assign(buffer);
					ShowElements(childwindows);

					int count = 0;
					childwindows->get_Length(&count);
					sprintf_s(buffer, "DumpAll: UICTL: '%-55s', window: %p: '%s', controls: %i\0", process.ImageName, process.WindowHandle, process.WindowTitle, count); trace(buffer);
					for (int i = 0; i < count; i++)
					{
						IUIAutomationElement *element = NULL;
						if (childwindows->GetElement(i, &element) == S_OK && element)
						{
							element->SetFocus();
							if (element->FindAll(TreeScope::TreeScope_Descendants, condition, &controls) == S_OK && controls)
							{
								sprintf_s(buffer, "DumpAll: UICTL: img: '%-55s', window: %p: '%s', controls: %i, idx: %i\0", process.ImageName, process.WindowHandle, process.WindowTitle, count, i);
								logprefix.assign(buffer);
								ShowElements(controls);
								
							} //if controls
							SAFE_REL_ARRAY(controls);
							
						} //if
						SAFE_REL(element);
					} //for
					SAFE_REL_ARRAY(childwindows);
					
				} //if
				else
				{
					sprintf_s(buffer, "DumpAll: WARNING: UIAUTOMATION NOT FOUND!! img: '%-55s', window: %p: '%s'\0", process.ImageName, process.WindowHandle, process.WindowTitle); log(buffer);
				} //else
			} //if
			
			VariantClear(&namevalue);
			SAFE_REL(namefind);
			SAFE_REL_ARRAY(childwindows);


			
		} //if window handle non null
	} //for

done:	
	SAFE_REL(namefind);
	SAFE_REL(window);
	SAFE_REL(element);
	SAFE_REL(condition);
	SAFE_REL_ARRAY(controls);
	SAFE_REL_ARRAY(childwindows);
	
}


void ListDescendants(IUIAutomation *ppAutomation, IUIAutomationElement* pParent, int indent)
{
	if (pParent == NULL) return;

	IUIAutomationTreeWalker* pControlWalker = NULL;
	IUIAutomationElement* pNode = NULL;

	ppAutomation->get_RawViewWalker(&pControlWalker);	if (pControlWalker == NULL) goto cleanup;
	pControlWalker->GetFirstChildElement(pParent, &pNode);	if (pNode == NULL) goto cleanup;

	while (pNode)
	{
		ShowElement(pNode);
		ListDescendants(ppAutomation, pNode, indent+1);
		IUIAutomationElement* pNext = NULL;
		pControlWalker->GetNextSiblingElement(pNode, &pNext);
		pNode->Release(); pNode = pNext;
	} //while

	cleanup:
		if (pControlWalker != NULL) pControlWalker->Release();

	if (pNode != NULL) pNode->Release();
}


void FortiClient(IUIAutomationElementArray *windows, IUIAutomationElement *root, IUIAutomation *ppAutomation)
{
	if (!fortclient) return;

	debug("FortiClient:-- START");
	// Find FortiClient. HWND
	RunningProcess process;
	if (!Find("FortiClient", process) || !process.WindowHandle)
	{
		debug("ProcessThread FortiClient not running...");
		return;
	} //if

	debug("ProcessThread FortiClient found..OK");
	ProcessFortiClient(windows, process.WindowHandle, ppAutomation);

}




void GlobalProtect(IUIAutomationElementArray *windows, IUIAutomationElement *root, IUIAutomation *ppAutomation)
{
	if (!globalprotect) return;

	char buffer[1024];
	HWND hWndTarget = 0;
	IUIAutomationElement *window = NULL;			// the window to automate
	IUIAutomationElement *taskbar = NULL;
	IUIAutomationElementArray *controls = NULL;		// window controls

	IUIAutomationCondition *condition = NULL;		// always true
	IUIAutomationCondition* namefind = NULL;

	IUIAutomationElement *button = NULL;			// button

	debug("GlobalProtect:-- START");

	if (ppAutomation->CreateTrueCondition(&condition) != S_OK || !condition)
	{
		debug("GlobalProtect unable to create TRUE condition (CreateTrueCondition) FAILED."); ERROR_DONE
	} //if

	// Find the task bar
	taskbar = FindElement(windows, "pane", "Taskbar");
	if (!taskbar)
	{
		debug("GlobalProtect Taskbar not found...");
		return;
	} //if

	{
		int count = 0;
		// list all controls.
		if (taskbar->FindAll(TreeScope::TreeScope_Descendants, condition, &controls) != S_OK || !controls)
		{
			debug("GlobalProtect FindAll window TreeScope_Descendants FAILED."); ERROR_DONE
		} //if

		debug("GlobalProtect Window Controls");
		controls->get_Length(&count);
		ShowElements(controls);

	} //scope

	// Find notification chevron (option 1), crappy option. have to clicky clicky, rather show the window directly.
	// Show all GA Windows (option 2) better option, find window by elements inside as multiple windows with the same name (GlobalProtect) and class with executable exist for PanGPA.exe
	// Look for a window that contains 'Not Connected' as a child window, then show that, and pass to UIAutomation.
	for (int i = 0; i < Processes.size(); i++)
	{
		RunningProcess& process = Processes[i];
		if (process.WindowHandle && strcmp(process.WindowTitle, "GlobalProtect") == 0)
		{
			if (!GetParent(process.WindowHandle))
			{
				sprintf_s(buffer, "\nChecking Window: hwnd: %p '%s'", process.WindowHandle, process.WindowTitle); debug(buffer);
				ChildElements.clear();
				logprefix = "GlobalProtect";
				EnumChildWindows(process.WindowHandle, EnumChildWindowCallback, NULL);	// win32 OLD, not used/required.

				// inspect
				for (int i = 0; i < (int)ChildElements.size(); i++)
				{
					RunningProcess child = ChildElements[i];
					if (child.WindowHandle)
					{
						if (strcmp(child.WindowTitle, "Not Connected") == 0)
						{
							debug("Showing window!!!");
							ShowWindow(process.WindowHandle, SW_SHOWNORMAL);
							hWndTarget = process.WindowHandle;
							break;
						} //if
					} //if
				} //for
			} //if
		} //if
	} //for

	if (hWndTarget)
	{
		// perhaps the window was not shown before, attempt to find
		VARIANT namevalue;
		VariantInit(&namevalue);
		namevalue.vt = VT_INT;
		namevalue.intVal = (int)hWndTarget;
		debug("looking for GlobalProtect");
		if (ppAutomation->CreatePropertyCondition(UIA_NativeWindowHandlePropertyId, namevalue, &namefind) == S_OK && namefind)
		{
			if (root->FindFirst(TreeScope::TreeScope_Children, namefind, &window) == S_OK && window)
			{
				debug("FOUND GlobalProtect");
			} //if
			else
			{
				debug("GlobalProtect NOT FOUND!!");
			} //else

		} //if
		VariantClear(&namevalue);
	} //if


	if (window)
	{
		int count = 0;

		// list all controls.
		if (window->FindAll(TreeScope::TreeScope_Descendants, condition, &controls) != S_OK || !controls)
		{
			debug("GlobalProtect FindAll window TreeScope_Descendants FAILED."); ERROR_DONE
		} //if

		debug("\nGlobalProtect Window Controls");
		controls->get_Length(&count);
		ShowElements(controls);

		// Find the Button
		button = FindElement(controls, "button", "Connect", TRUE);
		if (button)
		{
			debug("GlobalProtect Found Connect button...OK");				// we are not connected. check is active/clickable.
			IUIAutomationInvokePattern* iinvoke = NULL;
			if (button->GetCurrentPattern(UIA_InvokePatternId, reinterpret_cast<IUnknown**>(&iinvoke)) == S_OK && iinvoke)
			{
				iinvoke->Invoke();
			} //if
		} //if
		else
		{
			debug("GlobalProtect Connect button...Not found..");			// ignore and done for now.
		} //else

	} //if

done:
	SAFE_REL(namefind);
	SAFE_REL(window);
	SAFE_REL(condition);
	debug("GlobalProtect Releasing window controls."); SAFE_REL_ARRAY(controls);
}




void ProcessThread(void *parm)
{
	char buffer[1024];

	// Initialize COM
	IUIAutomation *ppAutomation = NULL;
	IUIAutomationElementArray *windows = NULL;		// all desktop windows
	IUIAutomationElement *root = NULL;				// desktop window
	IUIAutomationCondition *condition = NULL;		// always true

	if (CoInitializeEx(NULL, COINIT_MULTITHREADED) == S_OK)
	{
		debug("ProcessThread.CoInitializeEx.. OK");
	} //if
	else
	{
		debug("ProcessThread.CoInitializeEx.. NOT OK"); ERROR_DONE
	} //else

	// Obtain COM IID_IUIAutomation interface
	debug("ProcessThread.CoCreateInstance Obtain IID_IUIAutomation...");
	if (CoCreateInstance(CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER, IID_IUIAutomation, reinterpret_cast<void**>(&ppAutomation)) != S_OK || !ppAutomation)
	{
		debug("ProcessThread.CoCreateInstance IID_IUIAutomation FAILED..."); ERROR_DONE
	} //if
	debug("ProcessThread.CoCreateInstance Obtain IID_IUIAutomation...OK");

	debug("ProcessThread CreateTrueCondition");
	// Show all desktop windows
	if (ppAutomation->CreateTrueCondition(&condition) != S_OK || !condition)
	{
		debug("ProcessThread unable to create TRUE condition (CreateTrueCondition) FAILED."); ERROR_DONE
	} //if

	while (!quit)
	{
		debug("ProcessThread...");

		Processes.clear();
		EnumWindows(EnumWindowCallback, (LPARAM)&Processes);
		sprintf_s(buffer, "ProcessThread Enumerated %i windows.\0", (int)Processes.size()); debug(buffer);

		// IUI Automation
		if (ppAutomation->GetRootElement(&root) != S_OK || !root)		// gets the desktop
		{
			debug("ProcessThread unable to get root element (ppAutomation->GetRootElement) FAILED."); ERROR_DONE
		} //if

		// Show desktop name
		{
			BSTR bstr;
			root->get_CurrentName(&bstr);
			std::string cstr = utf8_encode(bstr);
			sprintf_s(buffer, "Root name: '%s'\0", cstr.c_str()); debug(buffer);
		} //scope

		// find all top level windows.
		debug("ProcessThread FindAll toplevel DesktopWindows");
		if (root->FindAll(TreeScope::TreeScope_Children, condition, &windows) != S_OK || !windows)
		{
			debug("ProcessThread FindAll toplevel DesktopWindows FAILED."); ERROR_DONE
		} //if

		// display all desktop windows
		ShowElements(windows);
		FortiClient(windows, root, ppAutomation);
		GlobalProtect(windows, root, ppAutomation);
		DumpAll(windows, root, ppAutomation);

		debug("Releasing Desktop windows."); SAFE_REL_ARRAY(windows);
		SAFE_REL(root);
		sprintf_s(buffer, "ProcessThread...waiting for %i\0", timer_in_ms); debug(buffer);
		if (timer_in_ms <= 1)
		{
			quit = true;
			break; //while
		}
		Sleep(timer_in_ms);
		//quit = true;

	} //while

	if (ppAutomation->GetRootElement(&root) != S_OK || !root)
	{
		debug("ProcessThread (2) unable to get root element (ppAutomation->GetRootElement) FAILED."); ERROR_DONE
	} //if
	
	debug("ProcessThread (2) FindAll toplevel DesktopWindows");
	if (root->FindAll(TreeScope::TreeScope_Children, condition, &windows) != S_OK || !windows)
	{
		debug("ProcessThread (2) FindAll toplevel DesktopWindows FAILED."); ERROR_DONE
	} //if
	
	if (fortclient)
	{
		DisconnectFortiClient(windows, root, ppAutomation);
	} //if


done:
	debug("Releasing Desktop windows."); SAFE_REL_ARRAY(windows);
	SAFE_REL(condition);
	SAFE_REL(root);

	// release COM IID_IUIAutomation interface
	SAFE_REL(ppAutomation);

	debug("ProcessThread:-- DONE");
	done = true;

	// release COM
	CoUninitialize();
	_endthread();

}


void DisconnectFortiClient(IUIAutomationElementArray *windows, IUIAutomationElement *root, IUIAutomation *ppAutomation)
{
	debug("DisconnectFortiClient:-- START");

	char buffer[2048];
	IUIAutomationElement *window = NULL;			// the window to automate
	IUIAutomationElementArray *controls = NULL;		// window controls
	IUIAutomationElement *button = NULL;			// button

	IUIAutomationCondition* namefind = NULL;
	IUIAutomationCondition *condition = NULL;		// always true

	VARIANT namevalue;
	VariantInit(&namevalue);
	
	RunningProcess process;
	if (!Find("FortiClient", process) || !process.WindowHandle)
	{
		debug("DisconnectFortiClient FortiClient not running...");
		return;
	} //if

	if (ppAutomation->CreateTrueCondition(&condition) != S_OK || !condition)
	{
		debug("DisconnectFortiClient unable to create TRUE condition (CreateTrueCondition) FAILED."); ERROR_DONE
	} //if
	

	// Find UIWindow
	namevalue.vt = VT_INT;
	namevalue.intVal = (int)process.WindowHandle;
	sprintf_s(buffer, "DisconnectFortiClient: looking for img: '%-55s', window: %p: '%s'\0", process.ImageName, process.WindowHandle, process.WindowTitle); trace(buffer);
	if (ppAutomation->CreatePropertyCondition(UIA_NativeWindowHandlePropertyId, namevalue, &namefind) != S_OK || !namefind)
	{
		debug("DisconnectFortiClient create condition namefind FAILED."); ERROR_DONE
	} //if

	if (root->FindFirst(TreeScope::TreeScope_Children, namefind, &window) == S_OK && window)
	{
		debug("FOUND FortiClient");
		// focus
		window->SetFocus();

		// find disconnect
		if (window->FindAll(TreeScope::TreeScope_Descendants, condition, &controls) != S_OK || !controls)
		{
			debug("DisconnectFortiClient FindAll window TreeScope_Descendants FAILED."); ERROR_DONE
		} //if

		debug("\nDisconnectFortiClient Window Controls");
		ShowElements(controls);

		button = FindElement(controls, "button", "Disconnect", TRUE);
		if (button)
		{
			debug("DisconnectFortiClient Found Disconnect button...OK");
			IUIAutomationInvokePattern* iinvoke = NULL;
			if (button->GetCurrentPattern(UIA_InvokePatternId, reinterpret_cast<IUnknown**>(&iinvoke)) == S_OK && iinvoke)
			{
				iinvoke->Invoke();
			} //if
		} //if
		else
		{
			debug("DisconnectFortiClient Connect button...Not found..");
		} //else
	} //if
	
	
done:
	VariantClear(&namevalue);
	
	SAFE_REL(namefind);
	SAFE_REL(button);
	SAFE_REL(window);
	SAFE_REL(condition);
	debug("DisconnectFortiClient Releasing window controls."); SAFE_REL_ARRAY(controls);
	
	debug("DisconnectFortiClient:-- DONE");
}

int WINAPI WinMain(HINSTANCE inst, HINSTANCE prev, LPSTR cmd, int show)
{
	log("WinMain:-- START");
	LPWSTR *args;
	int argc = 0;
	char buffer[1024];
	std::string prevarg = "";

	if (CoInitializeEx(NULL, COINIT_MULTITHREADED) == S_OK)
	{
		debug("WinMain.CoInitializeEx.. OK");
	} //if
	else
	{
		debug("WinMain.CoInitializeEx.. NOT OK"); ERROR_DONE
	} //else

	args = CommandLineToArgvW(GetCommandLineW(), &argc);
	for (int i = 0; i < argc; i++)
	{
		std::string str = utf8_encode(args[i]);
		sprintf_s(buffer, "Arg: #%i: '%s'\0", i, str.c_str()); log(buffer);
		if (strcmp(str.c_str(), "-debug") == 0) verbose = true;
		if (strcmp(str.c_str(), "-trace") == 0) veryverbose = true;
		if (strcmp(str.c_str(), "-visible") == 0) visiblewindowsonly = true;

		if (strcmp(prevarg.c_str(), "-timer") == 0)
		{
			timer_in_ms = atoi(str.c_str());
			timer_in_ms = clamp(timer_in_ms, 0, 60000 * 5);
			sprintf_s(buffer, "\t timer_in_ms: %i\0", timer_in_ms); debug(buffer);			
		}
		if (strcmp(prevarg.c_str(), "-image") == 0)
		{
			imagematch.assign(str.c_str());
		}
		
		if (strcmp(str.c_str(), "-fortclient") == 0) fortclient = true;
		if (strcmp(str.c_str(), "-globalprotect") == 0) globalprotect = true;
		
		// dump and show all windows, and all controls?
		if (strcmp(str.c_str(), "-all") == 0) dumpall = true;
		prevarg = str;
	} //for

	_beginthread(ProcessThread, 0, NULL);
	while (!quit)
	{
		if (FileExists("scripts/keepconnected.stop"))
		{
			quit = true;
			break;
		} //if

		Sleep(100);
	} //while

	while (!done)
	{
		Sleep(100);
	} //while

	Sleep(500);
done:
	log("WinMain DONE");
	CoUninitialize();

	return 0;
}
