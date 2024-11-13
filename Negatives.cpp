#include <windows.h>
#include <commctrl.h>
#include <gdiplus.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <shobjidl.h>
#include <string>
#include <fstream>
#include <sstream>
#include <map>
#include <vector>
#include <ctime>
#include <thread>
#include <algorithm>
#include <unordered_map>
#include <cmath>
#include <regex>
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "UxTheme.lib")
#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shcore.lib")
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

using namespace Gdiplus;
using namespace std;

#define MAX_PATH_LENGTH 260

enum {
    ID_LABEL_TITLE = 101, ID_LABEL_SUGGESTED, ID_BUTTON_BROWSE, ID_EDIT_PATH,
    ID_BUTTON_PROCESS, ID_BUTTON_PRINT, ID_BUTTON_GENERATE, ID_BUTTON_CLEAR,
    ID_LISTVIEW_STATUS, ID_PROGRESS_BAR, ID_STATUS_BAR,
    ID_COMBO_DEPARTMENTS, ID_MENU_FILE_EXIT, ID_MENU_HELP_ABOUT
};

struct Department {
    wstring name; int negativeSKUs; COLORREF color; vector<wstring> data;
    Department() : name(L""), negativeSKUs(0), color(RGB(0, 0, 0)) {}
    Department(const wstring& n, int c, COLORREF col) : name(n), negativeSKUs(c), color(col) {}
};

class AppData {
public:
    HWND hEditPath, hListViewStatus, hProgressBar, hStatusBar, hComboDepartments, hButtonGenerateCSVs, hButtonClear, hToolTip;
    HMENU hMenu; int processingCompleted; vector<Department> departments; ULONG_PTR gdiplusToken;
    wstring debugLogPath, tempDirectory; HFONT hFont, hBoldFont; unordered_map<wstring, wstring> departmentMapping;
    vector<wstring> tooltipTexts;
    AppData() : hEditPath(NULL), hListViewStatus(NULL), hProgressBar(NULL), hStatusBar(NULL),
        hComboDepartments(NULL), hButtonGenerateCSVs(NULL), hButtonClear(NULL),
        hToolTip(NULL), hMenu(NULL), processingCompleted(0), gdiplusToken(0),
        debugLogPath(L""), tempDirectory(L""), hFont(NULL), hBoldFont(NULL) {
    }
};

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void InitializeGUI(HWND, AppData*);
HICON GetStockIcon(SHSTOCKICONID, UINT);
void AddToolTip(HWND, HWND, const wchar_t*, AppData*);
void LogMessage(AppData*, const wstring&);
wstring Trim(const wstring&);
vector<wstring> ParseCSVLine(const wstring&);
wstring HTMLEscape(const wstring&);
void ShowErrorMessage(HWND, const wchar_t*);
void AddItemToListView(HWND, const wchar_t*, int);
void ClearListView(HWND);
void GetMostRecentFile(wchar_t*);
void BrowseForFile(HWND, AppData*);
void AutoProcessFileIfRecent(HWND, AppData*, const wchar_t*);
void ProcessNegativesFile(HWND, AppData*);
void GenerateCSVFiles(HWND, AppData*);
void ConvertCSVsToHTML(HWND, AppData*, const wstring&);
void OpenFolder(const wchar_t*);
void OpenHTMLInDefaultBrowser(const wchar_t*);
void ClearApplicationData(AppData*);
void ShowAboutDialog(HWND);

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int) {
    GdiplusStartupInput gdiplusStartupInput; AppData* pAppData = new AppData();
    if (GdiplusStartup(&pAppData->gdiplusToken, &gdiplusStartupInput, NULL) != Ok) {
        MessageBox(NULL, L"GDI+ Initialization Failed.", L"Error", MB_ICONERROR);
        delete pAppData; return -1;
    }
    INITCOMMONCONTROLSEX icex = { sizeof(icex), ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES | ICC_PROGRESS_CLASS | ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icex);
    WNDCLASSEX wc = {}; wc.cbSize = sizeof(WNDCLASSEX); wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc; wc.cbWndExtra = sizeof(AppData*);
    wc.hInstance = hInstance; wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = CreateSolidBrush(RGB(240, 240, 240)); wc.lpszClassName = L"NegativesProcessor";
    wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
    if (!RegisterClassEx(&wc)) {
        MessageBox(NULL, L"Window Registration Failed!", L"Error", MB_ICONERROR);
        GdiplusShutdown(pAppData->gdiplusToken); delete pAppData; return -1;
    }
    HWND hwnd = CreateWindowEx(0, wc.lpszClassName, L"", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 800, 600, NULL, NULL, hInstance, NULL);
    if (!hwnd) {
        MessageBox(NULL, L"Window Creation Failed!", L"Error", MB_ICONERROR);
        GdiplusShutdown(pAppData->gdiplusToken); delete pAppData; return -1;
    }
    ShowWindow(hwnd, SW_SHOW); UpdateWindow(hwnd);
    MSG msg; while (GetMessage(&msg, NULL, 0, 0)) { TranslateMessage(&msg); DispatchMessage(&msg); }
    GdiplusShutdown(pAppData->gdiplusToken); delete pAppData; return (int)msg.wParam;
}

HICON GetStockIcon(SHSTOCKICONID siid, UINT flags) {
    SHSTOCKICONINFO sii = { sizeof(sii) };
    if (SUCCEEDED(SHGetStockIconInfo(siid, SHGSI_ICON | flags, &sii))) return sii.hIcon;
    return NULL;
}

void InitializeGUI(HWND hwnd, AppData* pAppData) {
    pAppData->hFont = CreateFont(-16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    pAppData->hBoldFont = CreateFont(-18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    pAppData->hToolTip = CreateWindowEx(NULL, TOOLTIPS_CLASS, NULL, WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, hwnd, NULL, NULL, NULL);
    SetWindowPos(pAppData->hToolTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    HMENU hMenuBar = CreateMenu(); HMENU hFileMenu = CreatePopupMenu(); HMENU hHelpMenu = CreatePopupMenu();
    AppendMenu(hFileMenu, MF_STRING, ID_MENU_FILE_EXIT, L"Exit\tAlt+F4"); AppendMenu(hHelpMenu, MF_STRING, ID_MENU_HELP_ABOUT, L"About");
    AppendMenu(hMenuBar, MF_POPUP, (UINT_PTR)hFileMenu, L"&File"); AppendMenu(hMenuBar, MF_POPUP, (UINT_PTR)hHelpMenu, L"&Help");
    SetMenu(hwnd, hMenuBar);

    HWND hGroupBoxFile = CreateWindow(L"BUTTON", L"", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 15, 5, 760, 130, hwnd, NULL, NULL, NULL);
    SendMessage(hGroupBoxFile, WM_SETFONT, (WPARAM)pAppData->hBoldFont, TRUE);
    HWND hLabelTitle = CreateWindow(L"STATIC", L"Negatives Tool | Version 1.0 | Build 00021 11/14", WS_CHILD | WS_VISIBLE, 25, 15, 740, 25, hwnd, (HMENU)ID_LABEL_TITLE, NULL, NULL);
    SendMessage(hLabelTitle, WM_SETFONT, (WPARAM)pAppData->hFont, TRUE);
    HWND hLabelSuggested = CreateWindow(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, 25, 45, 740, 40, hwnd, (HMENU)ID_LABEL_SUGGESTED, NULL, NULL);
    SendMessage(hLabelSuggested, WM_SETFONT, (WPARAM)pAppData->hFont, TRUE);
    pAppData->hEditPath = CreateWindow(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT | ES_READONLY, 25, 95, 425, 25, hwnd, (HMENU)ID_EDIT_PATH, NULL, NULL);
    SendMessage(pAppData->hEditPath, WM_SETFONT, (WPARAM)pAppData->hFont, TRUE);
    HWND hButtonBrowse = CreateWindow(L"BUTTON", L" Browse", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_ICON, 465, 95, 75, 25, hwnd, (HMENU)ID_BUTTON_BROWSE, NULL, NULL);
    SendMessage(hButtonBrowse, BM_SETIMAGE, IMAGE_ICON, (LPARAM)GetStockIcon(SIID_FOLDER, SHGSI_SMALLICON));
    AddToolTip(hwnd, hButtonBrowse, L"Click to browse for a CSV file", pAppData);
    
    HWND hGroupBoxActions = CreateWindow(L"BUTTON", L"Actions", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 15, 140, 760, 160, hwnd, NULL, NULL, NULL);
    SendMessage(hGroupBoxActions, WM_SETFONT, (WPARAM)pAppData->hBoldFont, TRUE);
    HWND hButtonProcess = CreateWindow(L"BUTTON", L" Process File", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_ICON, 25, 160, 140, 35, hwnd, (HMENU)ID_BUTTON_PROCESS, NULL, NULL);
    SendMessage(hButtonProcess, BM_SETIMAGE, IMAGE_ICON, (LPARAM)GetStockIcon(SIID_APPLICATION, SHGSI_SMALLICON));
    AddToolTip(hwnd, hButtonProcess, L"Process the selected CSV file (Ctrl+P)", pAppData);
    pAppData->hButtonGenerateCSVs = CreateWindow(L"BUTTON", L" Generate CSVs", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_ICON, 180, 160, 140, 35, hwnd, (HMENU)ID_BUTTON_GENERATE, NULL, NULL);
    SendMessage(pAppData->hButtonGenerateCSVs, BM_SETIMAGE, IMAGE_ICON, (LPARAM)GetStockIcon(SIID_DOCNOASSOC, SHGSI_SMALLICON));
    AddToolTip(hwnd, pAppData->hButtonGenerateCSVs, L"Generate CSV files for departments", pAppData);
    HWND hButtonPrint = CreateWindow(L"BUTTON", L" Print", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_ICON, 335, 160, 140, 35, hwnd, (HMENU)ID_BUTTON_PRINT, NULL, NULL);
    SendMessage(hButtonPrint, BM_SETIMAGE, IMAGE_ICON, (LPARAM)GetStockIcon(SIID_PRINTER, SHGSI_SMALLICON));
    AddToolTip(hwnd, hButtonPrint, L"Generate printable HTML reports (Ctrl+R)", pAppData);
    pAppData->hButtonClear = CreateWindow(L"BUTTON", L" Clear", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_ICON, 490, 160, 110, 35, hwnd, (HMENU)ID_BUTTON_CLEAR, NULL, NULL);
    SendMessage(pAppData->hButtonClear, BM_SETIMAGE, IMAGE_ICON, (LPARAM)GetStockIcon(SIID_DELETE, SHGSI_SMALLICON));
    AddToolTip(hwnd, pAppData->hButtonClear, L"Clear all data and reset application", pAppData);
    pAppData->hComboDepartments = CreateWindow(WC_COMBOBOX, NULL, CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 25, 220, 575, 25, hwnd, (HMENU)ID_COMBO_DEPARTMENTS, NULL, NULL);
    SendMessage(pAppData->hComboDepartments, WM_SETFONT, (WPARAM)pAppData->hFont, TRUE);
    AddToolTip(hwnd, pAppData->hComboDepartments, L"Select department to print", pAppData);
    HWND hLabelStatus = CreateWindow(L"STATIC", L"Processing Results:", WS_CHILD | WS_VISIBLE, 25, 270, 550, 25, hwnd, NULL, NULL, NULL);
    SendMessage(hLabelStatus, WM_SETFONT, (WPARAM)pAppData->hBoldFont, TRUE);
    pAppData->hListViewStatus = CreateWindow(WC_LISTVIEW, L"", WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL, 25, 300, 600, 220, hwnd, (HMENU)ID_LISTVIEW_STATUS, NULL, NULL);
    SendMessage(pAppData->hListViewStatus, WM_SETFONT, (WPARAM)pAppData->hFont, TRUE);
    ListView_SetExtendedListViewStyle(pAppData->hListViewStatus, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
    LVCOLUMN lvc = {}; lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM; lvc.pszText = (LPWSTR)L"Major Department"; lvc.cx = 370; ListView_InsertColumn(pAppData->hListViewStatus, 0, &lvc);
    lvc.pszText = (LPWSTR)L"Negative SKUs Count"; lvc.cx = 210; ListView_InsertColumn(pAppData->hListViewStatus, 1, &lvc);
    pAppData->hProgressBar = CreateWindowEx(0, PROGRESS_CLASS, NULL, WS_CHILD | WS_VISIBLE | PBS_SMOOTH, 25, 530, 600, 20, hwnd, (HMENU)ID_PROGRESS_BAR, NULL, NULL);
    SendMessage(pAppData->hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
    SendMessage(pAppData->hProgressBar, PBM_SETPOS, 0, 0);
    pAppData->hStatusBar = CreateWindowEx(0, STATUSCLASSNAME, NULL, WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hwnd, (HMENU)ID_STATUS_BAR, NULL, NULL);
    int statusWidths[] = { 5, -1 }; SendMessage(pAppData->hStatusBar, SB_SETPARTS, 1, (LPARAM)statusWidths);
    SendMessage(pAppData->hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Ready");
    time_t now = time(NULL); tm tm_info; localtime_s(&tm_info, &now);
    wchar_t timeText[100]; wcsftime(timeText, 100, L"%Y-%m-%d %H:%M:%S", &tm_info);
    SendMessage(pAppData->hStatusBar, SB_SETTEXT, 1, (LPARAM)timeText);

    pAppData->departmentMapping = {
        {L"fencing & handling systems", L"Automotive"}, {L"livestock", L"Farm"}, {L"Small Animal", L"Farm"},
        {L"Dog", L"Pets"}, {L"Cat", L"Pets"}, {L"Birds & Critters", L"Pets"}, {L"toys", L"Seasonal"},
        {L"Backyard Leisure", L"Seasonal"}, {L"patio furniture & decor", L"Seasonal"}, {L"pest control", L"Seasonal"},
        {L"gardening", L"Seasonal"}, {L"lawn & garden tools", L"Seasonal"}, {L"lawn care", L"Seasonal"},
        {L"heating", L"Seasonal"}, {L"cooling", L"Seasonal"}, {L"consumables", L"Consumables"},
        {L"health & beauty", L"Consumables"}, {L"cleaning aids", L"Consumables"}, {L"Mens Accessories", L"Apparel"},
        {L"Mens Apparel", L"Apparel"}, {L"Womens Accessories", L"Apparel"}, {L"Womens Apparel", L"Apparel"},
        {L"Childrens Apparel", L"Apparel"}, {L"Womens Footwear", L"Apparel"}, {L"Mens Footwear", L"Apparel"},
        {L"hunting & archery", L"Sporting Goods"}, {L"ammunition", L"Sporting Goods"}, {L"Fluids & Chemicals", L"Automotive"},
        {L"Batteries and Accessories", L"Automotive"}, {L"Truck Tool Boxes", L"Automotive"}, {L"Cargo Management", L"Automotive"},
        {L"Tires, Towing, and Hitches", L"Automotive"}, {L"Small Engine Parts", L"Automotive"}, {L"hardware", L"Hardware"},
        {L"tools", L"Hardware"}, {L"plumbing", L"Hardware"}, {L"garage & shop equipment", L"Hardware"},
        {L"electrical", L"Hardware"}, {L"painting", L"Hardware"}, {L"Wood Cutting Equipment", L"Hardware"},
        {L"security; safety; locks", L"Hardware"}, {L"furnishings", L"Hardware"}
    };
    wchar_t suggestedFile[MAX_PATH_LENGTH] = L""; GetMostRecentFile(suggestedFile);
    if (wcslen(suggestedFile) > 0) {
        SetWindowText(pAppData->hEditPath, suggestedFile);
        HWND hLabelSuggested = GetDlgItem(hwnd, ID_LABEL_SUGGESTED);
        SetWindowText(hLabelSuggested, L"We found a recently downloaded negatives file. To use another file, press 'Browse'.");
        AutoProcessFileIfRecent(hwnd, pAppData, suggestedFile);
    }
    else {
        HWND hLabelSuggested = GetDlgItem(hwnd, ID_LABEL_SUGGESTED);
        SetWindowText(hLabelSuggested, L"No recently downloaded negatives file.");
    }
}

void AddToolTip(HWND hwndParent, HWND hwndControl, const wchar_t* text, AppData* pAppData) {
    pAppData->tooltipTexts.emplace_back(text);
    TOOLINFO toolInfo = { sizeof(toolInfo) };
    toolInfo.uFlags = TTF_SUBCLASS | TTF_IDISHWND;
    toolInfo.hwnd = hwndParent;
    toolInfo.uId = (UINT_PTR)hwndControl;
    toolInfo.lpszText = const_cast<LPWSTR>(pAppData->tooltipTexts.back().c_str());
    SendMessage(pAppData->hToolTip, TTM_ADDTOOL, 0, (LPARAM)&toolInfo);
}

void LogMessage(AppData* pAppData, const wstring& message) {
    if (pAppData->debugLogPath.empty()) return;
    wofstream logFile(pAppData->debugLogPath, ios::app);
    if (!logFile.is_open()) return;
    time_t now = time(NULL); tm tm_info; localtime_s(&tm_info, &now);
    wchar_t timeBuffer[30]; wcsftime(timeBuffer, 30, L"%Y-%m-%d %H:%M:%S", &tm_info);
    logFile << L"[" << timeBuffer << L"] " << message << L"\n";
}

wstring Trim(const wstring& str) {
    size_t first = str.find_first_not_of(L" \t\r\n");
    if (first == wstring::npos) return L"";
    size_t last = str.find_last_not_of(L" \t\r\n");
    return str.substr(first, last - first + 1);
}

vector<wstring> ParseCSVLine(const wstring& line) {
    vector<wstring> fields; wstringstream ss(line); wstring field; bool inQuotes = false; wchar_t ch;
    while (ss.get(ch)) {
        if (ch == L'\"') {
            if (inQuotes && ss.peek() == L'\"') { field += ch; ss.get(ch); }
            else { inQuotes = !inQuotes; }
        }
        else if (ch == L',' && !inQuotes) { fields.push_back(field); field.clear(); }
        else { field += ch; }
    }
    fields.push_back(field);
    return fields;
}

wstring HTMLEscape(const wstring& text) {
    wstring escaped; for (wchar_t ch : text) { switch (ch) { case L'&': escaped += L"&amp;"; break; case L'<': escaped += L"&lt;"; break; case L'>': escaped += L"&gt;"; break; case L'\"': escaped += L"&quot;"; break; case L'\'': escaped += L"&apos;"; break; default: escaped += ch; break; } }
    return escaped;
}

void ShowErrorMessage(HWND hwnd, const wchar_t* message) {
    MessageBox(hwnd, message, L"Error", MB_ICONERROR | MB_OK);
}

void AddItemToListView(HWND hListView, const wchar_t* department, int negativeSKUs) {
    LVITEM lvi = { 0 }; lvi.mask = LVIF_TEXT; lvi.pszText = const_cast<LPWSTR>(department);
    lvi.iItem = ListView_GetItemCount(hListView); ListView_InsertItem(hListView, &lvi);
    wchar_t countStr[10]; swprintf_s(countStr, 10, L"%d", negativeSKUs);
    ListView_SetItemText(hListView, lvi.iItem, 1, countStr);
}

void ClearListView(HWND hListView) {
    ListView_DeleteAllItems(hListView);
}

#include <windows.h>
#include <string>
using namespace std;

#define MAX_PATH_LENGTH MAX_PATH

#include <windows.h>
#include <string>
#include <regex>
using namespace std;

#define MAX_PATH_LENGTH MAX_PATH

#include <windows.h>
#include <string>
#include <regex>
using namespace std;

#define MAX_PATH_LENGTH MAX_PATH

#include <windows.h>
#include <string>
#include <regex>
using namespace std;

#define MAX_PATH_LENGTH MAX_PATH

void GetMostRecentFile(wchar_t* suggestedFile) {
    suggestedFile[0] = L'\0';

    wchar_t* userProfile = NULL;
    size_t len = 0;

    if (_wdupenv_s(&userProfile, &len, L"USERPROFILE") != 0 || userProfile == NULL) {
        return;
    }

    wchar_t paths[2][MAX_PATH_LENGTH];
    swprintf_s(paths[0], MAX_PATH_LENGTH, L"%s\\Downloads", userProfile);
    swprintf_s(paths[1], MAX_PATH_LENGTH, L"%s\\Desktop", userProfile);
    free(userProfile);

    WIN32_FIND_DATAW findFileData;
    HANDLE hFind;
    FILETIME latestTime = { 0 };
    wstring latestFile;

    // Regex pattern to match files starting with "export_" followed by date in the form YYYY-MM-DD HH_MM_SS.csv
    wregex exportPattern(LR"(export_\d{4}-\d{2}-\d{2} \d{2}_\d{2}_\d{2}\.csv)");

    for (int i = 0; i < 2; i++) {
        wchar_t searchPath[MAX_PATH_LENGTH];
        swprintf_s(searchPath, MAX_PATH_LENGTH, L"%s\\*.csv", paths[i]);

        hFind = FindFirstFileW(searchPath, &findFileData);
        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                if (!(findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
                    // Ensure the file matches the specific "export_" pattern
                    if (regex_match(findFileData.cFileName, exportPattern)) {
                        // Update if this file is the most recent
                        if (CompareFileTime(&findFileData.ftLastWriteTime, &latestTime) > 0) {
                            latestTime = findFileData.ftLastWriteTime;
                            latestFile = wstring(paths[i]) + L"\\" + findFileData.cFileName;
                        }
                    }
                }
            } while (FindNextFileW(hFind, &findFileData));
            FindClose(hFind);
        }
    }

    if (!latestFile.empty()) {
        wcsncpy_s(suggestedFile, MAX_PATH_LENGTH, latestFile.c_str(), _TRUNCATE);
    }
}




void BrowseForFile(HWND hwnd, AppData* pAppData) {
    OPENFILENAMEW ofn = {}; wchar_t szFileName[MAX_PATH_LENGTH] = L"";
    ofn.lStructSize = sizeof(OPENFILENAMEW); ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = L"CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0"; ofn.lpstrFile = szFileName;
    ofn.nMaxFile = MAX_PATH_LENGTH; ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST; ofn.lpstrDefExt = L"csv";
    if (GetOpenFileNameW(&ofn)) { SetWindowText(pAppData->hEditPath, szFileName); pAppData->processingCompleted = 0; ClearApplicationData(pAppData); }
}

void AutoProcessFileIfRecent(HWND hwnd, AppData* pAppData, const wchar_t* filePath) {
    WIN32_FILE_ATTRIBUTE_DATA fileInfo; if (GetFileAttributesEx(filePath, GetFileExInfoStandard, &fileInfo)) {
        FILETIME currentTime; GetSystemTimeAsFileTime(&currentTime);
        ULARGE_INTEGER current, fileTime; current.LowPart = currentTime.dwLowDateTime;
        current.HighPart = currentTime.dwHighDateTime; fileTime.LowPart = fileInfo.ftLastWriteTime.dwLowDateTime;
        fileTime.HighPart = fileInfo.ftLastWriteTime.dwHighDateTime;
        if (current.QuadPart < fileTime.QuadPart) return;
        double diffSeconds = (double)(current.QuadPart - fileTime.QuadPart) / 10000000.0;
        double diffHours = diffSeconds / 3600.0; if (diffHours <= 8.0) { thread(ProcessNegativesFile, hwnd, pAppData).detach(); }
    }
}

void ProcessNegativesFile(HWND hwnd, AppData* pAppData) {
    wchar_t filePath[MAX_PATH_LENGTH]; GetWindowText(pAppData->hEditPath, filePath, MAX_PATH_LENGTH);
    if (wcslen(filePath) == 0 || GetFileAttributesW(filePath) == INVALID_FILE_ATTRIBUTES) { ShowErrorMessage(hwnd, L"Please select a valid CSV file."); LogMessage(pAppData, L"Invalid CSV file selected."); return; }
    SendMessage(pAppData->hProgressBar, PBM_SETPOS, 10, 0); LogMessage(pAppData, L"Started processing CSV file.");
    wchar_t tempPath[MAX_PATH_LENGTH]; GetTempPath(MAX_PATH_LENGTH, tempPath);
    GUID guid; HRESULT hr = CoCreateGuid(&guid); if (FAILED(hr)) { ShowErrorMessage(hwnd, L"Failed to create GUID for temporary directory."); LogMessage(pAppData, L"CoCreateGuid failed."); return; }
    wchar_t guidString[40]; swprintf_s(guidString, 40, L"%08lX%04X%04X%04X%04X%08lX", guid.Data1, guid.Data2, guid.Data3,
        (guid.Data4[0] << 8) | guid.Data4[1],
        (guid.Data4[2] << 8) | guid.Data4[3],
        (unsigned long)((guid.Data4[4] << 24) | (guid.Data4[5] << 16) | (guid.Data4[6] << 8) | guid.Data4[7]));
    pAppData->tempDirectory = wstring(tempPath) + guidString;
    if (!CreateDirectory(pAppData->tempDirectory.c_str(), NULL)) {
        ShowErrorMessage(hwnd, L"Failed to create temporary directory."); LogMessage(pAppData, L"Failed to create temporary directory."); return;
    }
    pAppData->debugLogPath = pAppData->tempDirectory + L"\\debug.log"; LogMessage(pAppData, L"Temporary directory initialized.");
    map<wstring, vector<wstring>> majorDepartmentData; wifstream csvFile(filePath);
    if (!csvFile.is_open()) { ShowErrorMessage(hwnd, L"Failed to open CSV file."); LogMessage(pAppData, L"Failed to open CSV file."); return; }
    wstring line; getline(csvFile, line); int totalLines = 0;
    while (getline(csvFile, line)) {
        auto fields = ParseCSVLine(line); if (fields.size() < 10) { LogMessage(pAppData, L"Invalid CSV line: insufficient fields."); continue; }
        wstring itemNumber = Trim(fields[0]), description = Trim(fields[1]), minorDept = Trim(fields[2]), qtyTotalStr = Trim(fields[4]);
        auto it = pAppData->departmentMapping.find(minorDept); if (it == pAppData->departmentMapping.end()) { LogMessage(pAppData, L"Unknown department: " + minorDept); continue; }
        wstring majorDept = it->second; qtyTotalStr.erase(remove(qtyTotalStr.begin(), qtyTotalStr.end(), L','), qtyTotalStr.end());
        wchar_t* endPtr = nullptr; long quantity = wcstol(qtyTotalStr.c_str(), &endPtr, 10);
        if (endPtr == qtyTotalStr.c_str() || *endPtr != L'\0') { LogMessage(pAppData, L"Invalid quantity value in CSV line: " + qtyTotalStr); continue; }
        if (quantity < 0) { wstring formattedLine = L"\"" + itemNumber + L"\",\"" + description + L"\"," + to_wstring(quantity) + L",\"\""; majorDepartmentData[majorDept].push_back(formattedLine); }
        totalLines++;
    }
    csvFile.close(); LogMessage(pAppData, L"CSV file processing completed."); LogMessage(pAppData, L"Total lines processed: " + to_wstring(totalLines));
    SendMessage(pAppData->hProgressBar, PBM_SETPOS, 50, 0); pAppData->departments.clear();
    COLORREF colors[] = { RGB(255,99,132), RGB(54,162,235), RGB(255,206,86), RGB(75,192,192), RGB(153,102,255), RGB(255,159,64), RGB(199,199,199), RGB(83,102,255), RGB(255, 159, 28), RGB(0, 128, 0) };
    int colorIndex = 0; int totalNegatives = 0;
    for (auto& md : majorDepartmentData) totalNegatives += static_cast<int>(md.second.size());
    for (auto& md : majorDepartmentData) {
        int count = static_cast<int>(md.second.size());
        Department d(md.first, count, colors[colorIndex++ % (sizeof(colors) / sizeof(colors[0]))]);
        d.data = md.second;
        pAppData->departments.push_back(d);
        AddItemToListView(pAppData->hListViewStatus, d.name.c_str(), d.negativeSKUs);
        LogMessage(pAppData, L"Department added: " + d.name + L" with " + to_wstring(d.negativeSKUs) + L" negative SKUs.");
    }
    SendMessage(pAppData->hProgressBar, PBM_SETPOS, 80, 0); // Removed InvalidateRect for bar chart
    SendMessage(pAppData->hProgressBar, PBM_SETPOS, 100, 0); SendMessage(pAppData->hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Processing completed.");
    pAppData->processingCompleted = 1; SendMessage(pAppData->hProgressBar, PBM_SETPOS, 0, 0);
    LogMessage(pAppData, L"Processing completed successfully.");
    ClearListView(pAppData->hListViewStatus);
    for (const auto& dept : pAppData->departments) AddItemToListView(pAppData->hListViewStatus, dept.name.c_str(), dept.negativeSKUs);
    int totalNegativesCount = 0; for (const auto& dept : pAppData->departments) totalNegativesCount += dept.negativeSKUs;
    LVITEM lvi = { 0 }; lvi.mask = LVIF_TEXT; lvi.pszText = const_cast<LPWSTR>(L"Total Negatives");
    lvi.iItem = ListView_GetItemCount(pAppData->hListViewStatus); ListView_InsertItem(pAppData->hListViewStatus, &lvi);
    wchar_t totalStr[10]; swprintf_s(totalStr, 10, L"%d", totalNegativesCount); ListView_SetItemText(pAppData->hListViewStatus, lvi.iItem, 1, totalStr);
    SendMessage(pAppData->hComboDepartments, CB_RESETCONTENT, 0, 0); SendMessage(pAppData->hComboDepartments, CB_ADDSTRING, 0, (LPARAM)L"All Departments");
    for (const auto& dept : pAppData->departments) SendMessage(pAppData->hComboDepartments, CB_ADDSTRING, 0, (LPARAM)dept.name.c_str());
    SendMessage(pAppData->hComboDepartments, CB_SETCURSEL, 0, 0);
}

void GenerateCSVFiles(HWND hwnd, AppData* pAppData) {
    if (!pAppData->processingCompleted) { ShowErrorMessage(hwnd, L"Please process a file first."); return; }
    wstring csvFolder = pAppData->tempDirectory + L"\\CSV_Files";
    if (!CreateDirectory(csvFolder.c_str(), NULL) && GetLastError() != ERROR_ALREADY_EXISTS) {
        ShowErrorMessage(hwnd, L"Failed to create CSV_Files directory."); LogMessage(pAppData, L"Failed to create CSV_Files directory."); return;
    }
    for (const auto& dept : pAppData->departments) {
        wstring deptFilePath = csvFolder + L"\\" + dept.name + L".csv"; wofstream deptFile(deptFilePath);
        if (deptFile.is_open()) { deptFile << L"Item Number,Description,SOH,Your count\n"; for (const auto& dt : dept.data) deptFile << dt << L"\n"; deptFile.close(); LogMessage(pAppData, L"Department CSV written: " + deptFilePath); }
        else { ShowErrorMessage(hwnd, (wstring(L"Failed to write to ") + deptFilePath).c_str()); LogMessage(pAppData, L"Failed to write to department CSV: " + deptFilePath); }
    }
    OpenFolder(csvFolder.c_str());
}

void ConvertCSVsToHTML(HWND hwnd, AppData* pAppData, const wstring& selectedDepartment) {
    if (!pAppData->processingCompleted) { ShowErrorMessage(hwnd, L"Please process a file first."); return; }
    wchar_t htmlPath[MAX_PATH_LENGTH]; swprintf_s(htmlPath, MAX_PATH_LENGTH, L"%s\\Departments_Negatives.html", pAppData->tempDirectory.c_str());
    FILE* htmlFile = nullptr; if (_wfopen_s(&htmlFile, htmlPath, L"w, ccs=UTF-8") != 0 || !htmlFile) { ShowErrorMessage(hwnd, L"Failed to create HTML file."); LogMessage(pAppData, L"Failed to create HTML file."); return; }
    fwprintf(htmlFile, L"<html><head><meta charset=\"UTF-8\"><style>body { font-family: Arial; margin: 20px; }h1 { text-align: center; }table { width: 100%%; border-collapse: collapse; margin-top: 10px; }th, td { border: 1px solid #000; padding: 8px; text-align: left; }th { background-color: #f2f2f2; }.page-break { page-break-after: always; }</style><script>window.onload = function() { window.print(); };</script></head><body>");
    LogMessage(pAppData, L"HTML file creation started."); bool departmentFound = false;
    for (const auto& dept : pAppData->departments) {
        if (selectedDepartment != L"All Departments" && dept.name != selectedDepartment) continue;
        departmentFound = true; fwprintf(htmlFile, L"<h1>%s Negative SKUs</h1><table><tr><th>Item Number</th><th>Description</th><th>SOH</th><th>Your count</th></tr>", HTMLEscape(dept.name).c_str());
        int rowCount = 0; for (const auto& line : dept.data) {
            auto fields = ParseCSVLine(line); if (fields.size() < 4) continue;
            fwprintf(htmlFile, L"<tr>"); for (int i = 0; i < 4; ++i) fwprintf(htmlFile, L"<td>%s</td>", HTMLEscape(fields[i]).c_str());
            fwprintf(htmlFile, L"</tr>"); rowCount++;
        }
        fwprintf(htmlFile, L"</table><div class='page-break'></div>"); LogMessage(pAppData, L"HTML table created for department: " + dept.name + L" with " + to_wstring(rowCount) + L" items.");
        if (selectedDepartment != L"All Departments") break;
    }
    if (!departmentFound) { ShowErrorMessage(hwnd, L"No data available for the selected department."); fclose(htmlFile); _wremove(htmlPath); return; }
    fwprintf(htmlFile, L"</body></html>"); fclose(htmlFile); LogMessage(pAppData, L"HTML file creation completed: " + wstring(htmlPath)); OpenHTMLInDefaultBrowser(htmlPath);
}

void OpenFolder(const wchar_t* folderPath) {
    ShellExecute(NULL, L"open", folderPath, NULL, NULL, SW_SHOWDEFAULT);
}

void OpenHTMLInDefaultBrowser(const wchar_t* htmlPath) {
    ShellExecute(NULL, L"open", htmlPath, NULL, NULL, SW_SHOWDEFAULT);
}

void ClearApplicationData(AppData* pAppData) {
    if (!pAppData->tempDirectory.empty()) {
        SHFILEOPSTRUCT fileOp = { 0 }; fileOp.wFunc = FO_DELETE; wstring pathWithNull = pAppData->tempDirectory + L'\0';
        fileOp.pFrom = pathWithNull.c_str(); fileOp.fFlags = FOF_NO_UI; SHFileOperation(&fileOp); pAppData->tempDirectory.clear();
    }
    pAppData->departments.clear(); pAppData->processingCompleted = 0; ClearListView(pAppData->hListViewStatus);
    SendMessage(pAppData->hComboDepartments, CB_RESETCONTENT, 0, 0); SendMessage(pAppData->hProgressBar, PBM_SETPOS, 0, 0);
    SendMessage(pAppData->hStatusBar, SB_SETTEXT, 0, (LPARAM)L"Ready");
}

void ShowAboutDialog(HWND hwnd) {
    MessageBox(hwnd, L"Negatives Tool\nNegatives Tool | Version 1.0\nDeveloped by Tyler", L"About", MB_OK | MB_ICONINFORMATION);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    AppData* pAppData = reinterpret_cast<AppData*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE:
    {
        pAppData = new AppData(); SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pAppData);
        InitializeGUI(hwnd, pAppData); LogMessage(pAppData, L"Application started."); SetTimer(hwnd, 1, 1000, NULL);
    }
    break;
    case WM_COMMAND:
    {
        if (pAppData == nullptr) break; switch (LOWORD(wParam)) {
        case ID_BUTTON_BROWSE: BrowseForFile(hwnd, pAppData); break;
        case ID_BUTTON_PROCESS: thread(ProcessNegativesFile, hwnd, pAppData).detach(); break;
        case ID_BUTTON_GENERATE: thread(GenerateCSVFiles, hwnd, pAppData).detach(); break;
        case ID_BUTTON_PRINT:
        {
            int selIndex = (int)SendMessage(pAppData->hComboDepartments, CB_GETCURSEL, 0, 0); wchar_t selectedDept[256] = L"";
            SendMessage(pAppData->hComboDepartments, CB_GETLBTEXT, selIndex, (LPARAM)selectedDept); selectedDept[255] = L'\0';
            ConvertCSVsToHTML(hwnd, pAppData, selectedDept);
            break;
        }
        case ID_BUTTON_CLEAR: ClearApplicationData(pAppData); break;
        case ID_MENU_FILE_EXIT: PostMessage(hwnd, WM_CLOSE, 0, 0); break;
        case ID_MENU_HELP_ABOUT: ShowAboutDialog(hwnd); break;
        }
    }
    break;
    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam; SetTextColor(hdcStatic, RGB(0, 0, 0)); SetBkMode(hdcStatic, TRANSPARENT);
        return (INT_PTR)GetSysColorBrush(COLOR_WINDOW);
    }
    break;
    case WM_TIMER:
    {
        if (wParam == 1 && pAppData != nullptr) {
            time_t now = time(NULL); tm tm_info; localtime_s(&tm_info, &now); wchar_t timeText[100];
            wcsftime(timeText, 100, L"%Y-%m-%d %H:%M:%S", &tm_info);
            SendMessage(pAppData->hStatusBar, SB_SETTEXT, 1, (LPARAM)timeText);
        }
    }
    break;
    case WM_DESTROY:
    {
        if (pAppData != nullptr) {
            LogMessage(pAppData, L"Application exiting."); if (pAppData->hFont) DeleteObject(pAppData->hFont);
            if (pAppData->hBoldFont) DeleteObject(pAppData->hBoldFont); ClearApplicationData(pAppData);
            delete pAppData;
        }
        PostQuitMessage(0);
    }
    break;
    default: return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}
