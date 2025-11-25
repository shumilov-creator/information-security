#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <string>
#include <sstream>
#include <vector>
#include <fstream>
#include <random>
#include <cstdint>
#include <algorithm>
#include <iomanip>
#include <cwchar>
#include <shlwapi.h>
#include <locale>
#include <commctrl.h>
#include <wincrypt.h>
#include <bcrypt.h>
#include <uxtheme.h>
#include <cstring>

#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Bcrypt.lib")
#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "UxTheme.lib")
#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

using namespace std;

// ------------------------
// GUI handles / state
// ------------------------
HINSTANCE hInst;
HWND hEditKey, hEditInput, hEditOutput, hBtnSaveStatus, hComboMode, hBtnCopy, hBtnOpenOutputDir, hBtnHelp, hStaticKeyLength;
HWND hStaticCurrentUser, hEditCurrentUser;
HWND hExCombo, hExSend, hExRefresh, hExInList, hExOutList, hExOpenFolder;
HFONT hFont, hFontTitle;
HWND hTooltip;

wstring BASE_DIR_PATH = L"";
wstring g_CurrentUser = L"";
const wstring KEYS_FILENAME = L"Ключи.txt";
const wstring PASSWORDS_FILENAME = L"Пароли.txt";

wstring g_LastEncryptedFile = L"";
wstring g_LastUser = L"";
wstring g_LastInputText = L"";
int     g_FileCounter = 0;

static vector<wstring> g_AllRecipients;

// ------------------------
// Цвета оставлены по умолчанию системными значениями

const int UI_H = 30;
const int UI_PAD = 12;
const int UI_GAP = 10;
const int UI_BTN = 160;

// ------------------------
// Styling helpers
// ------------------------
int ScaleByDpi(HWND hWnd, int px) {
    UINT dpi = 96;
    if (HMODULE hU = GetModuleHandleW(L"user32")) {
        typedef UINT(WINAPI* FN)(HWND);
        static FN get = (FN)GetProcAddress(hU, "GetDpiForWindow");
        if (get) dpi = get(hWnd);
    }
    return MulDiv(px, (int)dpi, 96);
}

RECT GetChildRect(HWND parent, HWND child) {
    RECT rc{ 0,0,0,0 };
    if (!child) return rc;
    GetWindowRect(child, &rc);
    MapWindowPoints(NULL, parent, (POINT*)&rc, 2);
    return rc;
}

RECT UnionRects(const RECT& a, const RECT& b) {
    RECT r{};
    r.left = min(a.left, b.left);
    r.top = min(a.top, b.top);
    r.right = max(a.right, b.right);
    r.bottom = max(a.bottom, b.bottom);
    return r;
}

RECT CombineRects(HWND parent, std::initializer_list<HWND> children, int padding) {
    RECT combined{ 0,0,0,0 };
    bool has = false;
    for (HWND c : children) {
        if (!c) continue;
        RECT rc = GetChildRect(parent, c);
        if (!has) { combined = rc; has = true; }
        else combined = UnionRects(combined, rc);
    }
    if (!has) { SetRectEmpty(&combined); return combined; }
    InflateRect(&combined, padding, padding);
    return combined;
}

void PaintGradientBackground(HDC hdc, const RECT& rc) {
    FillRect(hdc, &rc, GetSysColorBrush(COLOR_WINDOW));
}

void DrawSoftCard(HWND, HDC, const RECT&) {
    // Дополнительное декоративное оформление отключено
}

void ApplyExplorerTheme(HWND hCtrl) {
    if (hCtrl) SetWindowTheme(hCtrl, L"Explorer", NULL);
}

void ApplyEditPadding(HWND hCtrl, HWND hWnd) {
    if (!hCtrl) return;
    int pad = ScaleByDpi(hWnd, 8);
    SendMessage(hCtrl, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(pad, pad));
}

const wchar_t* BTN_VARIANT_PROP = L"BTN_VARIANT";

HWND CreateStyledButton(HWND hWnd, int controlId, const wchar_t* text, int x, int y, int w, int h, bool primary) {
    HWND btn = CreateWindowW(L"BUTTON", text,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        x, y, w, h,
        hWnd, (HMENU)controlId, hInst, NULL);
    if (btn) {
        SetPropW(btn, BTN_VARIANT_PROP, (HANDLE)(primary ? 1 : 2));
    }
    return btn;
}

bool DrawStyledButton(LPDRAWITEMSTRUCT dis) {
    if (!dis || dis->CtlType != ODT_BUTTON) return false;

    ULONG_PTR variant = (ULONG_PTR)GetPropW(dis->hwndItem, BTN_VARIANT_PROP);
    if (variant != 1 && variant != 2) return false;

    bool primary = variant == 1;
    bool pressed = (dis->itemState & ODS_SELECTED) != 0;
    bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    bool focused = (dis->itemState & ODS_FOCUS) != 0;

    COLORREF base = primary ? RGB(64, 120, 255) : RGB(240, 243, 250);
    COLORREF border = primary ? RGB(36, 90, 210) : RGB(210, 214, 226);
    if (pressed) {
        base = primary ? RGB(52, 101, 224) : RGB(226, 229, 238);
        border = primary ? RGB(30, 76, 190) : RGB(195, 199, 210);
    }
    if (disabled) {
        base = RGB(230, 230, 230);
        border = RGB(200, 200, 200);
    }

    RECT rc = dis->rcItem;
    HBRUSH br = CreateSolidBrush(base);
    HPEN pen = CreatePen(PS_SOLID, 1, border);
    HGDIOBJ oldPen = SelectObject(dis->hDC, pen);
    HGDIOBJ oldBrush = SelectObject(dis->hDC, br);

    RoundRect(dis->hDC, rc.left, rc.top, rc.right, rc.bottom, 10, 10);

    SelectObject(dis->hDC, oldPen);
    SelectObject(dis->hDC, oldBrush);
    DeleteObject(pen);
    DeleteObject(br);

    int len = GetWindowTextLengthW(dis->hwndItem);
    wstring caption(len + 1, L'\0');
    GetWindowTextW(dis->hwndItem, caption.data(), len + 1);
    caption.resize(len);

    SetBkMode(dis->hDC, TRANSPARENT);
    COLORREF textColor = primary ? RGB(255, 255, 255) : RGB(40, 46, 60);
    if (disabled) textColor = RGB(150, 150, 150);
    SetTextColor(dis->hDC, textColor);

    RECT textRc = rc;
    textRc.left += 12; textRc.right -= 12; textRc.top += 2; textRc.bottom -= 2;
    if (pressed) OffsetRect(&textRc, 1, 1);
    DrawTextW(dis->hDC, caption.c_str(), -1, &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    if (focused) {
        RECT focusRc = rc; InflateRect(&focusRc, -4, -4);
        DrawFocusRect(dis->hDC, &focusRc);
    }
    return true;
}

// ------------------------
// IDs
// ------------------------
#define ID_BTN_OPEN_KEYS_FILE      107
#define ID_BTN_CLEAR               109
#define ID_BTN_GENERATE_KEY        106
#define ID_BTN_ENCRYPT             3
#define ID_BTN_DECRYPT             4
#define ID_BTN_DECRYPT_LAST        114
#define ID_BTN_IN_FILE             1
#define ID_BTN_COPY                110
#define ID_EDIT_KEY                100
#define ID_EDIT_INPUT              101
#define ID_EDIT_OUTPUT             102
#define ID_BTN_SAVE_STATUS         113
#define ID_BTN_OPEN_OUTPUT_DIR     115
#define ID_BTN_HELP                116
#define ID_STATIC_KEY_LENGTH       117
#define ID_LOGIN_EDIT_USER         201
#define ID_LOGIN_EDIT_PASSWORD     202
#define ID_LOGIN_OK                203
#define ID_LOGIN_CANCEL            204
#define ID_BTN_CHANGE_USER         118
#define ID_LOGIN_BTN_NEW_USER      205
#define ID_STATIC_CURRENT_USER     119
#define ID_EDIT_CURRENT_USER       120
#define ID_LOGIN_SHOWPASS          206
// Exchange block
#define ID_EXCH_COMBO              402
#define ID_EXCH_SEND               403
#define ID_EXCH_REFRESH            404
#define ID_EXCH_IN_LIST            405
#define ID_EXCH_OPEN_FOLDER        406
#define ID_EXCH_OUT_LIST           407

// Keys manager (ListView)
#define ID_KEYS_LISTVIEW     551
#define ID_KEYS_ADD_CURRENT  552
#define ID_KEYS_DELETE       553
#define ID_KEYS_COPY         554
#define ID_KEYS_REFRESH      555
#define ID_KEYS_SET_CURRENT  556
#define ID_KEYS_RENAME       557
#define ID_KEYS_EXPORT       558
#define ID_KEYS_IMPORT       559

// Menu IDs
#define IDM_FILE_OPEN        600
#define IDM_FILE_SAVE        601
#define IDM_FILE_EXIT        602
#define IDM_TOOLS_KEYS       610
#define IDM_TOOLS_EXCHANGE   611

// Prototypes
void HandleEncryptDecrypt(HWND, bool, bool);
ATOM RegisterKeysClass(HINSTANCE);
ATOM RegisterExchangeClass(HINSTANCE);
void OpenKeysManager(HWND);
void OpenExchangeWindow(HWND);
ATOM MyRegisterClass(HINSTANCE);
ATOM RegisterLoginClass(HINSTANCE);
BOOL InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK LoginWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK KeysWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK ExchangeWndProc(HWND, UINT, WPARAM, LPARAM);

// ------------------------
// S-boxes GOST 28147-89 ("Тест")
// ------------------------
const uint8_t S[8][16] = {
    {4,10,9,2,13,8,0,14,6,11,1,12,7,15,5,3},
    {14,11,4,12,6,13,15,10,2,3,8,1,0,7,5,9},
    {5,8,1,13,10,3,4,2,14,15,12,7,6,0,9,11},
    {7,13,10,1,0,8,9,15,14,4,6,12,11,2,5,3},
    {6,12,7,1,5,15,13,8,4,10,9,14,0,3,11,2},
    {4,11,10,0,7,2,1,13,3,6,8,5,9,12,15,14},
    {13,11,4,1,3,15,5,9,0,10,14,7,6,8,2,12},
    {1,15,13,0,5,7,10,4,9,2,3,14,6,11,8,12}
};

// ------------------------
// Paths & helpers
// ------------------------
wstring GetExePath() {
    wchar_t p[MAX_PATH];
    GetModuleFileNameW(NULL, p, MAX_PATH);
    wstring s(p);
    return s.substr(0, s.find_last_of(L"\\/"));
}
wstring GetOutputDir() { return BASE_DIR_PATH + L"\\Выходные файлы"; }
wstring GetExchangeDir() { return BASE_DIR_PATH + L"\\Обмен"; }

wstring TrimWString(const wstring& s) {
    size_t a = s.find_first_not_of(L" \t\r\n");
    if (a == wstring::npos) return L"";
    size_t b = s.find_last_not_of(L" \t\r\n");
    return s.substr(a, b - a + 1);
}
bool EnsureDirectoryExists(const wstring& path) {
    if (PathIsDirectoryW(path.c_str())) return true;
    return CreateDirectoryW(path.c_str(), NULL) || GetLastError() == ERROR_ALREADY_EXISTS;
}
string WStringToUTF8(const wstring& w) {
    if (w.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), NULL, 0, NULL, NULL);
    string s(n, 0);
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), &s[0], n, NULL, NULL);
    return s;
}
wstring UTF8ToWString(const string& s) {
    if (s.empty()) return {};
    size_t st = 0;
    if (s.size() >= 3 && (uint8_t)s[0] == 0xEF && (uint8_t)s[1] == 0xBB && (uint8_t)s[2] == 0xBF) st = 3;
    string v = s.substr(st);
    int n = MultiByteToWideChar(CP_UTF8, 0, v.c_str(), (int)v.size(), NULL, 0);
    wstring w(n, 0);
    MultiByteToWideChar(CP_UTF8, 0, v.c_str(), (int)v.size(), &w[0], n);
    return w;
}

// Base64
string Base64Encode(const vector<uint8_t>& d) {
    static const char t[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    string o; int val = 0, valb = -6;
    for (uint8_t c : d) { val = (val << 8) + c; valb += 8; while (valb >= 0) { o.push_back(t[(val >> valb) & 0x3F]); valb -= 6; } }
    if (valb > -6) o.push_back(t[((val << 8) >> (valb + 8)) & 0x3F]);
    while (o.size() % 4) o.push_back('=');
    return o;
}
vector<uint8_t> Base64Decode(const string& s) {
    static const unsigned char T[256] = {
        64,64,64,64,64,64,64,64,64,64,64,64,64,64,64,64,
        64,64,64,64,64,64,64,64,64,64,64,64,64,64,64,64,
        64,64,64,64,64,64,64,64,64,64,64,62,64,64,64,63,
        52,53,54,55,56,57,58,59,60,61,64,64,64,64,64,64,
        64, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14,
        15,16,17,18,19,20,21,22,23,24,25,64,64,64,64,64,
        64,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,
        41,42,43,44,45,46,47,48,49,50,51
    };
    vector<uint8_t> o; int val = 0, valb = -8;
    for (unsigned char c : s) {
        if (c == '=' || T[c] == 64) break;
        val = (val << 6) + T[c]; valb += 6;
        if (valb >= 0) { o.push_back((val >> valb) & 0xFF); valb -= 8; }
    }
    return o;
}

// low-level
uint32_t ROL32(uint32_t v, int n) { return (v << n) | (v >> (32 - n)); }
uint32_t get_uint32(const uint8_t* p) { uint32_t v; memcpy(&v, p, sizeof(v)); return v; }
void put_uint32(uint8_t* p, uint32_t v) { memcpy(p, &v, sizeof(v)); }

// case-insensitive contains (Unicode)
static wchar_t ToLowerRu(wchar_t c) { return (wchar_t)towlower(c); }
static bool ICaseContains(const wstring& hay, const wstring& needle) {
    if (needle.empty()) return true;
    wstring h = hay, n = needle;
    transform(h.begin(), h.end(), h.begin(), ToLowerRu);
    transform(n.begin(), n.end(), n.begin(), ToLowerRu);
    return h.find(n) != wstring::npos;
}

// HEX key utils
wstring NormalizeHex64(const wstring& raw) {
    wstring s = TrimWString(raw), out; out.reserve(s.size());
    for (wchar_t ch : s) if (iswxdigit(ch)) out.push_back((wchar_t)towupper(ch));
    return out;
}
bool ValidateKey(const wstring& k) {
    if (k.size() != 64) return false;
    for (wchar_t c : k) if (!iswxdigit(c)) return false;
    return true;
}
vector<uint8_t> hexStringToBytes(const wstring& hex) {
    vector<uint8_t> b; if (hex.size() != 64) return b;
    for (size_t i = 0;i < hex.size();i += 2) { wchar_t p[3] = { hex[i],hex[i + 1],L'\0' }; b.push_back((uint8_t)wcstoul(p, nullptr, 16)); }
    return b;
}

// Users list from passwords file
vector<wstring> ListSavedUsers() {
    vector<wstring> users;
    wstring path = BASE_DIR_PATH + L"\\" + PASSWORDS_FILENAME;
    ifstream fin(path, ios::binary); if (!fin) return users;
    string line;
    while (getline(fin, line)) {
        line.erase(remove(line.begin(), line.end(), '\r'), line.end());
        if (line.empty()) continue;
        wstring wl = UTF8ToWString(line);
        size_t pos = wl.find(L" — ");
        if (pos != wstring::npos && pos > 0) {
            wstring u = TrimWString(wl.substr(0, pos));
            if (!u.empty()) users.push_back(u);
        }
    }
    sort(users.begin(), users.end());
    users.erase(unique(users.begin(), users.end()), users.end());
    return users;
}

// Key length indicator
void UpdateKeyLengthIndicator(HWND) {
    wchar_t buf[256]; GetWindowTextW(hEditKey, buf, 256);
    wstringstream w; w << wcslen(buf) << L"/64 символов";
    SetWindowTextW(hStaticKeyLength, w.str().c_str());
}

void CopyKeyToClipboard(HWND hWnd, const wstring& s) {
    if (!OpenClipboard(hWnd)) return;
    EmptyClipboard();
    SIZE_T sz = (s.size() + 1) * sizeof(wchar_t);
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE, sz);
    if (h) {
        void* p = GlobalLock(h);
        memcpy(p, s.c_str(), sz);
        GlobalUnlock(h);
        SetClipboardData(CF_UNICODETEXT, h);
    }
    CloseClipboard();
    MessageBoxW(hWnd, L"Ключ скопирован в буфер обмена!", L"Информация", MB_OK | MB_ICONINFORMATION);
}

// RNG key
wstring GenerateRandomKey() {
    uint8_t buf[32];
    if (BCryptGenRandom(NULL, buf, (ULONG)sizeof(buf), BCRYPT_USE_SYSTEM_PREFERRED_RNG) != 0) {
        random_device rd; for (auto& b : buf) b = (uint8_t)rd();
    }
    static const wchar_t HEX[] = L"0123456789ABCDEF";
    wstring out; out.resize(64);
    for (int i = 0;i < 32;++i) { out[i * 2] = HEX[(buf[i] >> 4) & 0x0F]; out[i * 2 + 1] = HEX[buf[i] & 0x0F]; }
    return out;
}

// open folder
bool OpenFolderInExplorer(const wstring& dir) {
    if (!EnsureDirectoryExists(dir)) return false;
    PIDLIST_ABSOLUTE pidl = nullptr;
    HRESULT hr = SHParseDisplayName(dir.c_str(), nullptr, &pidl, 0, nullptr);
    if (SUCCEEDED(hr) && pidl) {
        hr = SHOpenFolderAndSelectItems(pidl, 0, nullptr, 0);
        CoTaskMemFree(pidl);
        if (SUCCEEDED(hr)) return true;
    }
    HINSTANCE h = ShellExecuteW(nullptr, L"open", dir.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return (INT_PTR)h > 32;
}

// DPAPI
bool DPAPI_Protect(const vector<uint8_t>& in, vector<uint8_t>& out) {
    DATA_BLOB a{ (DWORD)in.size(), (BYTE*)in.data() }, b{};
    if (!CryptProtectData(&a, L"", NULL, NULL, NULL, 0, &b)) return false;
    out.assign(b.pbData, b.pbData + b.cbData);
    LocalFree(b.pbData);
    return true;
}
bool DPAPI_Unprotect(const vector<uint8_t>& in, vector<uint8_t>& out) {
    DATA_BLOB a{ (DWORD)in.size(), (BYTE*)in.data() }, b{};
    LPWSTR d = NULL;
    if (!CryptUnprotectData(&a, &d, NULL, NULL, NULL, 0, &b)) return false;
    if (d) LocalFree(d);
    out.assign(b.pbData, b.pbData + b.cbData);
    LocalFree(b.pbData);
    return true;
}

// Passwords save/load
void SaveUserPassword(const wstring& user, const wstring& pass) {
    if (user.empty() || pass.empty()) return;
    vector<uint8_t> plain((uint8_t*)pass.data(), (uint8_t*)pass.data() + pass.size() * sizeof(wchar_t)), prot;
    if (!DPAPI_Protect(plain, prot)) return;
    wstring path = BASE_DIR_PATH + L"\\" + PASSWORDS_FILENAME;
    EnsureDirectoryExists(BASE_DIR_PATH);
    ifstream fin(path, ios::binary);
    string line, buf;
    bool found = false; wstring prefix = user + L" — ";
    while (getline(fin, line)) {
        line.erase(remove(line.begin(), line.end(), '\r'), line.end());
        if (line.empty()) continue;
        wstring wl = UTF8ToWString(line);
        if (wl.rfind(prefix, 0) == 0) { buf += WStringToUTF8(prefix) + Base64Encode(prot) + "\n"; found = true; }
        else buf += line + "\n";
    }
    fin.close();
    if (!found) buf += WStringToUTF8(prefix) + Base64Encode(prot) + "\n";
    ofstream fout(path, ios::binary); if (fout) fout << buf;
}
wstring LoadUserPassword(const wstring& user) {
    if (user.empty()) return L"";
    wstring path = BASE_DIR_PATH + L"\\" + PASSWORDS_FILENAME;
    ifstream fin(path, ios::binary); if (!fin) return L"";
    string line; wstring prefix = user + L" — ";
    while (getline(fin, line)) {
        line.erase(remove(line.begin(), line.end(), '\r'), line.end());
        if (line.empty()) continue;
        wstring wl = UTF8ToWString(line);
        if (wl.rfind(prefix, 0) == 0) {
            string b = WStringToUTF8(wl.substr(prefix.length()));
            vector<uint8_t> blob = Base64Decode(b), plain;
            if (blob.empty() || !DPAPI_Unprotect(blob, plain) || plain.size() % sizeof(wchar_t)) return L"";
            return wstring((wchar_t*)plain.data(), plain.size() / sizeof(wchar_t));
        }
    }
    return L"";
}
bool VerifyPassword(const wstring& user, const wstring& pass) { wstring s = LoadUserPassword(user); return !s.empty() && s == pass; }
bool UserExists(const wstring& user) { return !LoadUserPassword(user).empty(); }

// ------------------------
// MULTI-KEY SUPPORT (label + current)
// ------------------------
struct UserKey {
    wstring key;      // 64HEX
    wstring when;     // "YYYY-MM-DD HH:MM" (optional)
    wstring label;    // optional
    bool    isCurrent = false;
};

static vector<wstring> SplitByEmDash(const wstring& s) {
    vector<wstring> out; size_t start = 0, pos; const wstring sep = L" — ";
    while ((pos = s.find(sep, start)) != wstring::npos) { out.push_back(s.substr(start, pos - start)); start = pos + sep.size(); }
    out.push_back(s.substr(start)); return out;
}
static wstring NowStamp() {
    SYSTEMTIME st; GetLocalTime(&st);
    wstringstream w; w << setfill(L'0') << setw(4) << st.wYear << L"-" << setw(2) << st.wMonth << L"-" << setw(2) << st.wDay
        << L" " << setw(2) << st.wHour << L":" << setw(2) << st.wMinute;
    return w.str();
}
static wstring SafeJoin(const vector<wstring>& parts) {
    wstring s; for (size_t i = 0;i < parts.size();++i) { if (i) s += L" — "; s += parts[i]; } return s;
}

static vector<UserKey> LoadAllUserKeys(const wstring& user) {
    vector<UserKey> list; if (user.empty()) return list;
    wstring path = BASE_DIR_PATH + L"\\" + KEYS_FILENAME;
    ifstream fin(path, ios::binary); if (!fin) return list;
    string line;
    while (getline(fin, line)) {
        line.erase(remove(line.begin(), line.end(), '\r'), line.end());
        if (line.empty()) continue;
        wstring wl = UTF8ToWString(line);
        auto parts = SplitByEmDash(wl);
        if (parts.size() >= 2 && TrimWString(parts[0]) == user) {
            UserKey uk; uk.key = NormalizeHex64(TrimWString(parts[1])); if (!ValidateKey(uk.key)) continue;
            if (parts.size() >= 3) uk.when = TrimWString(parts[2]);
            if (parts.size() >= 4) uk.label = TrimWString(parts[3]);
            if (parts.size() >= 5) uk.isCurrent = (TrimWString(parts[4]) == L"*current*");
            list.push_back(uk);
        }
    }
    return list;
}

wstring LoadUserKey(const wstring& user) {
    auto v = LoadAllUserKeys(user);
    if (v.empty()) return L"";
    for (auto& k : v) if (k.isCurrent) return k.key;
    return v.back().key;
}

void SaveUserKey(const wstring& user, const wstring& rawKey, const wstring& label = L"") {
    if (user.empty()) return;
    wstring key = NormalizeHex64(rawKey); if (!ValidateKey(key)) return;
    EnsureDirectoryExists(BASE_DIR_PATH);
    wstring path = BASE_DIR_PATH + L"\\" + KEYS_FILENAME;

    auto existing = LoadAllUserKeys(user);
    bool already = any_of(existing.begin(), existing.end(), [&](const UserKey& uk) { return uk.key == key; });
    if (already) return;

    ofstream fout(path, ios::app | ios::binary); if (!fout) return;
    vector<wstring> parts{ user, key, NowStamp() }; if (!label.empty()) parts.push_back(label);
    wstring line = SafeJoin(parts) + L"\n"; fout << WStringToUTF8(line);
}

bool SetCurrentUserKey(const wstring& user, const wstring& key) {
    if (user.empty() || !ValidateKey(key)) return false;
    wstring path = BASE_DIR_PATH + L"\\" + KEYS_FILENAME;
    ifstream fin(path, ios::binary); if (!fin) return false;
    string line, out;
    while (getline(fin, line)) {
        string raw = line; raw.erase(remove(raw.begin(), raw.end(), '\r'), raw.end());
        wstring wl = UTF8ToWString(raw);
        auto parts = SplitByEmDash(wl);
        if (parts.size() >= 2 && TrimWString(parts[0]) == user) {
            wstring curKey = NormalizeHex64(TrimWString(parts[1]));
            vector<wstring> rebuild;
            rebuild.push_back(TrimWString(parts[0]));
            rebuild.push_back(curKey);
            if (parts.size() >= 3) rebuild.push_back(TrimWString(parts[2]));
            if (parts.size() >= 4) rebuild.push_back(TrimWString(parts[3]));
            if (curKey == key) rebuild.push_back(L"*current*");
            out += WStringToUTF8(SafeJoin(rebuild)) + "\n";
        }
        else out += line + "\n";
    }
    fin.close(); ofstream fout(path, ios::binary); if (!fout) return false; fout << out; return true;
}

bool UpdateUserKeyLabel(const wstring& user, const wstring& key, const wstring& newLabel) {
    if (user.empty() || !ValidateKey(key)) return false;
    wstring path = BASE_DIR_PATH + L"\\" + KEYS_FILENAME;
    ifstream fin(path, ios::binary); if (!fin) return false;
    string line, out; bool changed = false;
    while (getline(fin, line)) {
        string raw = line; raw.erase(remove(raw.begin(), raw.end(), '\r'), raw.end());
        wstring wl = UTF8ToWString(raw);
        auto parts = SplitByEmDash(wl);
        if (parts.size() >= 2 && TrimWString(parts[0]) == user && NormalizeHex64(TrimWString(parts[1])) == key) {
            wstring when = parts.size() >= 3 ? TrimWString(parts[2]) : L"";
            bool isCur = (parts.size() >= 5 && TrimWString(parts[4]) == L"*current*");
            vector<wstring> rebuild{ user, key };
            if (!when.empty()) rebuild.push_back(when); else rebuild.push_back(NowStamp());
            if (!newLabel.empty()) rebuild.push_back(newLabel);
            if (isCur) rebuild.push_back(L"*current*");
            out += WStringToUTF8(SafeJoin(rebuild)) + "\n"; changed = true;
        }
        else out += line + "\n";
    }
    fin.close(); ofstream fout(path, ios::binary); if (!fout) return false; fout << out; return changed;
}

bool DeleteUserKey(const wstring& user, const wstring& key) {
    if (user.empty() || !ValidateKey(key)) return false;
    wstring path = BASE_DIR_PATH + L"\\" + KEYS_FILENAME;
    ifstream fin(path, ios::binary); if (!fin) return false;
    string line, out; bool removed = false;
    while (getline(fin, line)) {
        string raw = line; raw.erase(remove(raw.begin(), raw.end(), '\r'), raw.end());
        wstring wl = UTF8ToWString(raw);
        auto parts = SplitByEmDash(wl);
        if (parts.size() >= 2 && TrimWString(parts[0]) == user && NormalizeHex64(TrimWString(parts[1])) == key) {
            removed = true; continue;
        }
        out += line + "\n";
    }
    fin.close(); ofstream fout(path, ios::binary); if (!fout) return false; fout << out; return removed;
}

bool ExportUserKeyToFile(HWND hWnd, const wstring& key) {
    if (!ValidateKey(key)) return false;
    OPENFILENAMEW ofn = { sizeof(ofn) }; wchar_t file[MAX_PATH] = L"";
    ofn.hwndOwner = hWnd; ofn.lpstrFile = file; ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"Key file (*.key)\0*.key\0Text file (*.txt)\0*.txt\0All files (*.*)\0*.*\0";
    ofn.nFilterIndex = 1; ofn.Flags = OFN_OVERWRITEPROMPT; ofn.lpstrDefExt = L"key";
    if (!GetSaveFileNameW(&ofn)) return false;
    ofstream f(file, ios::binary); if (!f) return false; f << WStringToUTF8(key); return true;
}
bool ImportUserKeyFromFile(HWND hWnd, const wstring& user) {
    OPENFILENAMEW ofn = { sizeof(ofn) }; wchar_t file[MAX_PATH] = L"";
    ofn.hwndOwner = hWnd; ofn.lpstrFile = file; ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"Key/Text files (*.key;*.txt)\0*.key;*.txt\0All files (*.*)\0*.*\0";
    ofn.nFilterIndex = 1; ofn.Flags = OFN_FILEMUSTEXIST;
    if (!GetOpenFileNameW(&ofn)) return false;
    ifstream fin(wstring(file), ios::binary); if (!fin) return false;
    string s((istreambuf_iterator<char>(fin)), {}); wstring w = NormalizeHex64(UTF8ToWString(s));
    if (!ValidateKey(w)) return false;
    wstring label = wstring(file); size_t pos = label.find_last_of(L"\\/"); if (pos != wstring::npos) label = label.substr(pos + 1);
    SaveUserKey(user, w, label);
    return true;
}

// ------------------------
// GOST block cipher
// ------------------------
void GostEncryptBlock(const uint8_t* in, uint8_t* out, const vector<uint8_t>& k) {
    uint32_t n1 = get_uint32(in), n2 = get_uint32(in + 4);
    uint32_t kk[8]; for (int i = 0;i < 8;++i) kk[i] = get_uint32(&k[i * 4]);
    static const int ENC_K[32] = { 0,1,2,3,4,5,6,7, 0,1,2,3,4,5,6,7, 0,1,2,3,4,5,6,7, 7,6,5,4,3,2,1,0 };
    for (int r = 0;r < 32;++r) {
        uint32_t t = n1 + kk[ENC_K[r]], s = 0;
        for (int i = 0;i < 8;++i) s |= (uint32_t)S[i][(t >> (4 * i)) & 0x0F] << (4 * i);
        s = ROL32(s, 11) ^ n2;
        if (r < 31) { n2 = n1; n1 = s; }
        else n2 = s;
    }
    put_uint32(out, n1); put_uint32(out + 4, n2);
}
void GostDecryptBlock(const uint8_t* in, uint8_t* out, const vector<uint8_t>& k) {
    uint32_t n1 = get_uint32(in), n2 = get_uint32(in + 4);
    uint32_t kk[8]; for (int i = 0;i < 8;++i) kk[i] = get_uint32(&k[i * 4]);
    static const int DEC_K[32] = { 0,1,2,3,4,5,6,7, 7,6,5,4,3,2,1,0, 7,6,5,4,3,2,1,0, 7,6,5,4,3,2,1,0 };
    for (int r = 0;r < 32;++r) {
        uint32_t t = n1 + kk[DEC_K[r]], s = 0;
        for (int i = 0;i < 8;++i) s |= (uint32_t)S[i][(t >> (4 * i)) & 0x0F] << (4 * i);
        s = ROL32(s, 11) ^ n2;
        if (r < 31) { n2 = n1; n1 = s; }
        else n2 = s;
    }
    put_uint32(out, n1); put_uint32(out + 4, n2);
}

// ------------------------
// Container GOST0 + mode + IV + data
// ------------------------
vector<uint8_t> GostEncryptContainer(const wstring& text, const vector<uint8_t>& key, int mode) {
    string utf8 = WStringToUTF8(text);
    vector<uint8_t> data(utf8.begin(), utf8.end());
    uint8_t pad = 8 - (data.size() % 8); if (pad == 0) pad = 8; data.insert(data.end(), pad, pad);
    vector<uint8_t> out(data.size());

    uint8_t iv[8] = { 0 };
    if (mode == 1 || mode == 2) BCryptGenRandom(NULL, iv, sizeof(iv), BCRYPT_USE_SYSTEM_PREFERRED_RNG);

    uint8_t prev[8] = { 0 }, tmp[8];
    if (mode != 0) memcpy(prev, iv, 8);

    for (size_t i = 0;i < data.size();i += 8) {
        if (mode == 0) {
            GostEncryptBlock(&data[i], &out[i], key);
        }
        else if (mode == 1) { // CBC
            for (int j = 0;j < 8;++j) data[i + j] ^= prev[j];
            GostEncryptBlock(&data[i], &out[i], key);
            memcpy(prev, &out[i], 8);
        }
        else { // CFB
            GostEncryptBlock(prev, tmp, key);
            for (int j = 0;j < 8;++j) { out[i + j] = tmp[j] ^ data[i + j]; prev[j] = out[i + j]; }
        }
    }

    vector<uint8_t> cont;
    const uint8_t magic[5] = { 'G','O','S','T','0' };
    cont.insert(cont.end(), magic, magic + 5);
    cont.push_back((uint8_t)mode);
    cont.insert(cont.end(), iv, iv + 8);
    cont.insert(cont.end(), out.begin(), out.end());
    return cont;
}
wstring GostDecryptContainer(const vector<uint8_t>& bin, const vector<uint8_t>& key) {
    if (bin.size() < 5 + 1 + 8) return L"Ошибка: слишком короткие данные.";
    if (!(bin[0] == 'G' && bin[1] == 'O' && bin[2] == 'S' && bin[3] == 'T' && bin[4] == '0'))
        return L"Ошибка: неизвестный контейнер.";
    int mode = bin[5]; const uint8_t* iv = &bin[6]; size_t pos = 5 + 1 + 8;
    if ((bin.size() - pos) % 8 != 0) return L"Ошибка: размер некратен 8.";
    vector<uint8_t> enc(bin.begin() + pos, bin.end()), out(enc.size());
    uint8_t prev[8] = { 0 }, tmp[8]; memcpy(prev, iv, 8);

    for (size_t i = 0;i < enc.size();i += 8) {
        if (mode == 0) {
            GostDecryptBlock(&enc[i], &out[i], key);
        }
        else if (mode == 1) {
            memcpy(tmp, &enc[i], 8);
            GostDecryptBlock(&enc[i], &out[i], key);
            for (int j = 0;j < 8;++j) out[i + j] ^= prev[j];
            memcpy(prev, tmp, 8);
        }
        else {
            GostEncryptBlock(prev, tmp, key);
            for (int j = 0;j < 8;++j) { out[i + j] = tmp[j] ^ enc[i + j]; prev[j] = enc[i + j]; }
        }
    }

    if (!out.empty()) {
        uint8_t pad = out.back();
        if (pad == 0 || pad > 8 || out.size() < pad) return L"Ошибка: недопустимый паддинг.";
        for (size_t k = out.size() - pad;k < out.size();++k) if (out[k] != pad) return L"Ошибка: некорректный паддинг.";
        out.resize(out.size() - pad);
    }
    return UTF8ToWString(string(out.begin(), out.end()));
}
wstring DirectTextEncrypt(const wstring& text, const wstring& keyStr, int mode) {
    wstring n = NormalizeHex64(keyStr); vector<uint8_t> key = hexStringToBytes(n);
    if (key.size() != 32) return L"Ошибка: некорректный ключ. Нужны 64 HEX-символа.";
    vector<uint8_t> c = GostEncryptContainer(text, key, mode);
    return UTF8ToWString(Base64Encode(c));
}
wstring DirectTextDecrypt(const wstring& in, const wstring& keyStr, int) {
    wstring n = NormalizeHex64(keyStr); vector<uint8_t> key = hexStringToBytes(n);
    if (key.size() != 32) return L"Ошибка: некорректный ключ.";
    string b = WStringToUTF8(in); b.erase(remove_if(b.begin(), b.end(), ::isspace), b.end());
    while (b.size() % 4) b += '=';
    vector<uint8_t> bin = Base64Decode(b);
    if (bin.size() < 5 + 1 + 8) return L"Ошибка: повреждённые данные.";
    return GostDecryptContainer(bin, key);
}

// ------------------------
// Autosave + filenames
// ------------------------
wstring GenerateSimpleFileName(const wstring& user, bool enc, int counter) {
    wstring outDir = GetOutputDir(), userDir = outDir + L"\\" + user;
    EnsureDirectoryExists(outDir); EnsureDirectoryExists(userDir);
    SYSTEMTIME st; GetLocalTime(&st);
    wstringstream w;
    w << setfill(L'0') << setw(4) << st.wYear << L"-" << setw(2) << st.wMonth << L"-" << setw(2) << st.wDay << L"_"
        << setw(2) << st.wHour << setw(2) << st.wMinute;
    return userDir + L"\\" + (enc ? L"_Зашифрованный_" : L"_Расшифрованный_") + w.str() + L"_" + to_wstring(counter) + L".txt";
}
bool IsNewInputText(const wstring& t) {
    wstring s = TrimWString(t);
    if (g_LastInputText.empty() || s != g_LastInputText) { g_LastInputText = s; g_FileCounter++; return true; }
    return false;
}
bool AutoSaveResult(HWND hWnd, const wstring& text, const wstring& user, bool enc, int counter) {
    wstring fp = GenerateSimpleFileName(user, enc, counter);
    wstring userDir = GetOutputDir() + L"\\" + user;
    ofstream f(fp, ios::binary); if (!f) return false;
    if (enc) { g_LastEncryptedFile = fp; g_LastUser = user; f << WStringToUTF8(text); }
    else { f << (char)0xEF << (char)0xBB << (char)0xBF << WStringToUTF8(text); }
    f.close();
    wstring msg = L"Файл автосохранён.\n\nПуть: " + fp;
    if (MessageBoxW(hWnd, (msg + L"\n\nОткрыть папку?").c_str(), L"Автосохранение", MB_YESNO | MB_ICONINFORMATION) == IDYES)
        OpenFolderInExplorer(userDir);
    return true;
}
bool SaveResultToFile(const wstring& text, const wstring& path, bool enc) {
    ofstream f(path, ios::binary); if (!f) return false;
    if (enc) f << WStringToUTF8(text);
    else f << (char)0xEF << (char)0xBB << (char)0xBF << WStringToUTF8(text);
    f.close(); return true;
}

// ------------------------
// Encrypt/Decrypt handler
// ------------------------
void HandleEncryptDecrypt(HWND hWnd, bool encrypt, bool decryptLast) {
    wchar_t keyBuf[256], inputBuf[20000];
    GetWindowTextW(hEditKey, keyBuf, 256);
    GetWindowTextW(hEditInput, inputBuf, 20000);

    wstring keyStr = NormalizeHex64(keyBuf);
    wstring inputTxt = TrimWString(inputBuf);
    int mode = (int)SendMessage(hComboMode, CB_GETCURSEL, 0, 0);

    if (g_CurrentUser.empty()) { MessageBoxW(hWnd, L"Сначала войдите в систему.", L"Ошибка", MB_OK | MB_ICONERROR); return; }
    if (!ValidateKey(keyStr)) {
        MessageBoxW(hWnd, L"Ключ должен быть 64 HEX символа (0–9, A–F).", L"Ошибка", MB_OK | MB_ICONERROR); return;
    }

    if (!encrypt) {
        if (decryptLast) {
            if (g_LastEncryptedFile.empty() || g_LastUser != g_CurrentUser) {
                MessageBoxW(hWnd, L"Нет последнего зашифрованного файла для этого пользователя.", L"Ошибка", MB_OK | MB_ICONERROR); return;
            }
            ifstream fin(g_LastEncryptedFile, ios::binary);
            if (!fin) { MessageBoxW(hWnd, L"Не удалось открыть последний зашифрованный файл.", L"Ошибка", MB_OK | MB_ICONERROR); return; }
            string data((istreambuf_iterator<char>(fin)), {});
            inputTxt = UTF8ToWString(data); SetWindowTextW(hEditInput, (L"[🔒] " + inputTxt).c_str());
        }
        if (inputTxt.empty()) { MessageBoxW(hWnd, L"Введите зашифрованный текст или выберите файл.", L"Ошибка", MB_OK | MB_ICONERROR); return; }
    }
    else {
        if (inputTxt.empty()) { MessageBoxW(hWnd, L"Введите исходный текст или выберите файл.", L"Ошибка", MB_OK | MB_ICONERROR); return; }
    }

    wstring result;
    if (encrypt) {
        IsNewInputText(inputTxt);
        result = DirectTextEncrypt(inputTxt, keyStr, mode);
        if (result.rfind(L"Ошибка:", 0) == 0) { MessageBoxW(hWnd, result.c_str(), L"Ошибка шифрования", MB_OK | MB_ICONERROR); return; }
        AutoSaveResult(hWnd, result, g_CurrentUser, true, g_FileCounter);
        SetWindowTextW(hBtnSaveStatus, L"📥 Автосохранено (Зашифровано)");
    }
    else {
        result = DirectTextDecrypt(inputTxt, keyStr, mode);
        if (result.rfind(L"Ошибка:", 0) == 0) { MessageBoxW(hWnd, result.c_str(), L"Ошибка расшифрования", MB_OK | MB_ICONERROR); return; }
        AutoSaveResult(hWnd, result, g_CurrentUser, false, g_FileCounter);
        SetWindowTextW(hBtnSaveStatus, L"📥 Автосохранено (Расшифровано)");
    }

    SetWindowTextW(hEditOutput, result.c_str());
    // сохраняем ключ в список (если его ещё нет; метка пустая)
    SaveUserKey(g_CurrentUser, keyStr, L"");
}

// ------------------------
// File dialogs / tooltip
// ------------------------
void HandleOpenFile(HWND hWnd) {
    OPENFILENAMEW ofn = { sizeof(ofn) }; wchar_t file[MAX_PATH] = L"";
    ofn.hwndOwner = hWnd; ofn.lpstrFile = file; ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"Текстовые файлы (*.txt)\0*.txt\0Все файлы (*.*)\0*.*\0";
    ofn.nFilterIndex = 1; ofn.lpstrInitialDir = BASE_DIR_PATH.c_str();
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    if (GetOpenFileNameW(&ofn)) {
        ifstream in(wstring(file), ios::binary);
        if (!in) { MessageBoxW(hWnd, L"Не удалось открыть файл.", L"Ошибка", MB_OK | MB_ICONERROR); return; }
        string buf((istreambuf_iterator<char>(in)), {});
        SetWindowTextW(hEditInput, UTF8ToWString(buf).c_str());
    }
}
void HandleSaveResult(HWND hWnd) {
    wchar_t outBuf[20000]; GetWindowTextW(hEditOutput, outBuf, 20000); wstring res(outBuf);
    if (res.empty() || res.rfind(L"Ошибка:", 0) == 0) { MessageBoxW(hWnd, L"Поле результата пусто или содержит ошибку.", L"Ошибка", MB_OK | MB_ICONERROR); return; }

    OPENFILENAMEW ofn = { sizeof(ofn) }; wchar_t file[MAX_PATH] = L"";
    ofn.hwndOwner = hWnd; ofn.lpstrFile = file; ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"Текстовые файлы (*.txt)\0*.txt\0Все файлы (*.*)\0*.*\0";
    ofn.nFilterIndex = 1; ofn.lpstrInitialDir = BASE_DIR_PATH.c_str();
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT; ofn.lpstrDefExt = L"txt";
    if (GetSaveFileNameW(&ofn)) {
        bool isEnc = false; string p = WStringToUTF8(res), b = p; b.erase(remove_if(b.begin(), b.end(), ::isspace), b.end());
        vector<uint8_t> bin = Base64Decode(b);
        if (bin.size() >= 5 && bin[0] == 'G' && bin[1] == 'O' && bin[2] == 'S' && bin[3] == 'T' && bin[4] == '0') isEnc = true;

        if (SaveResultToFile(res, wstring(file), isEnc)) {
            SetWindowTextW(hBtnSaveStatus, L"📥 Сохранено вручную");
            MessageBoxW(hWnd, L"Файл успешно сохранён!", L"Успех", MB_OK | MB_ICONINFORMATION);
        }
        else {
            MessageBoxW(hWnd, L"Ошибка при сохранении файла.", L"Ошибка", MB_OK | MB_ICONERROR);
        }
    }
}
void InitToolTips(HWND hWnd) {
    hTooltip = CreateWindowExW(0, TOOLTIPS_CLASSW, NULL, WS_POPUP | TTS_ALWAYSTIP,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, hWnd, NULL, hInst, NULL);
    if (!hTooltip) return;
    TOOLINFOW ti = { sizeof(TOOLINFOW) }; ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS; ti.hwnd = hWnd; ti.uId = (UINT_PTR)hComboMode;
    ti.lpszText = (LPWSTR)L"Режимы ГОСТ 28147-89:\n• ECB — простой блочный режим.\n• CBC — сцепление блоков.\n• CFB — потоковый режим.\nКонтейнер: GOST0 + MODE + IV + данные.";
    SendMessageW(hTooltip, TTM_ADDTOOLW, 0, (LPARAM)&ti);
}

// ------------------------
// Exchange (helpers)
// ------------------------
void Exchange_RefreshInboxList() {
    if (!hExInList || g_CurrentUser.empty()) return;
    SendMessage(hExInList, LB_RESETCONTENT, 0, 0);
    wstring inbox = GetExchangeDir() + L"\\" + g_CurrentUser + L"\\Inbox";
    EnsureDirectoryExists(inbox);
    WIN32_FIND_DATAW fd; HANDLE h = FindFirstFileW((inbox + L"\\*.txt").c_str(), &fd);
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
                SendMessage(hExInList, LB_ADDSTRING, 0, (LPARAM)fd.cFileName);
        } while (FindNextFileW(h, &fd)); FindClose(h);
    }
}
void Exchange_RefreshOutboxList() {
    if (!hExOutList || g_CurrentUser.empty()) return;
    SendMessage(hExOutList, LB_RESETCONTENT, 0, 0);
    wstring outbox = GetExchangeDir() + L"\\" + g_CurrentUser + L"\\Outbox";
    EnsureDirectoryExists(outbox);
    WIN32_FIND_DATAW fd; HANDLE h = FindFirstFileW((outbox + L"\\*.txt").c_str(), &fd);
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
                SendMessage(hExOutList, LB_ADDSTRING, 0, (LPARAM)fd.cFileName);
        } while (FindNextFileW(h, &fd)); FindClose(h);
    }
}
static void Exchange_RebuildRecipientsCombo(const wstring& typed, bool userTyping) {
    if (!hExCombo) return;
    DWORD start = 0, end = 0; SendMessage(hExCombo, CB_GETEDITSEL, (WPARAM)&start, (LPARAM)&end);
    SendMessage(hExCombo, CB_RESETCONTENT, 0, 0);
    for (const auto& u : g_AllRecipients) if (typed.empty() || ICaseContains(u, typed))
        SendMessage(hExCombo, CB_ADDSTRING, 0, (LPARAM)u.c_str());
    SetWindowTextW(hExCombo, typed.c_str());
    DWORD len = (DWORD)typed.size(); if (start > len) start = len; if (end > len) end = len;
    SendMessage(hExCombo, CB_SETEDITSEL, 0, MAKELPARAM(start, end));
    if (userTyping) SendMessage(hExCombo, CB_SHOWDROPDOWN, TRUE, 0);
}
void Exchange_UpdateAllRecipients() {
    g_AllRecipients.clear();
    for (const auto& u : ListSavedUsers()) if (u != g_CurrentUser) g_AllRecipients.push_back(u);
    Exchange_RebuildRecipientsCombo(L"", false);
}

bool Exchange_ExtractKeyAndPayload(const wstring& fileData, wstring& keyOut, wstring& payloadOut) {
    size_t nl = fileData.find(L'\n');
    if (nl == wstring::npos) return false;

    wstring firstLine = fileData.substr(0, nl);
    firstLine.erase(remove(firstLine.begin(), firstLine.end(), L'\r'), firstLine.end());
    wstring trimmed = TrimWString(firstLine);

    wstring lower = trimmed;
    transform(lower.begin(), lower.end(), lower.begin(), ::towlower);
    if (lower.rfind(L"key:", 0) != 0) return false;

    wstring candidate = NormalizeHex64(trimmed.substr(4));
    if (!ValidateKey(candidate)) return false;

    wstring rest = fileData.substr(nl + 1);
    payloadOut = TrimWString(rest);
    keyOut = candidate;
    return !payloadOut.empty();
}

bool Exchange_SendTo(const wstring& recipient, const wstring& textToSend, int mode) {
    if (recipient.empty()) return false;
    if (!UserExists(recipient)) return false;
    wstring recKey = LoadUserKey(recipient);
    if (!ValidateKey(recKey)) return false;

    vector<uint8_t> key = hexStringToBytes(recKey);
    vector<uint8_t> cont = GostEncryptContainer(textToSend, key, mode);
    wstring b64 = UTF8ToWString(Base64Encode(cont));

    wstring fileContent = L"KEY:" + recKey + L"\r\n" + b64;

    wstring base = GetExchangeDir();
    wstring inRec = base + L"\\" + recipient + L"\\Inbox";
    wstring outMe = base + L"\\" + g_CurrentUser + L"\\Outbox";
    EnsureDirectoryExists(base); EnsureDirectoryExists(base + L"\\" + recipient);
    EnsureDirectoryExists(base + L"\\" + g_CurrentUser); EnsureDirectoryExists(inRec); EnsureDirectoryExists(outMe);

    SYSTEMTIME st; GetLocalTime(&st);
    wstringstream ts; ts << setfill(L'0') << setw(4) << st.wYear << L"-" << setw(2) << st.wMonth << L"-" << setw(2) << st.wDay << L"_"
        << setw(2) << st.wHour << setw(2) << st.wMinute;
    wstring fname = L"from_" + g_CurrentUser + L"_to_" + recipient + L"_" + ts.str() + L"_" + to_wstring(++g_FileCounter) + L".txt";

    auto writeFile = [&](const wstring& dir) { ofstream f(dir + L"\\" + fname, ios::binary); if (!f) return false; f << WStringToUTF8(fileContent); return true; };
    bool ok1 = writeFile(inRec), ok2 = writeFile(outMe);
    return ok1 && ok2;
}

// ------------------------
// User change
// ------------------------
void ChangeUser(HWND hWnd) {
    wchar_t kb[256]; GetWindowTextW(hEditKey, kb, 256);
    wstring cur = NormalizeHex64(kb);
    if (!g_CurrentUser.empty() && ValidateKey(cur)) SaveUserKey(g_CurrentUser, cur, L"");
    HWND dlg = CreateWindowW(L"LOGIN_GUI", L"Смена пользователя",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 560, 420, hWnd, NULL, hInst, NULL);
    if (dlg) {
        ShowWindow(dlg, SW_SHOW); UpdateWindow(dlg);
        MSG m;
        while (IsWindow(dlg) && GetMessage(&m, NULL, 0, 0)) {
            if (m.message == WM_USER + 1) {
                if (hStaticCurrentUser && !g_CurrentUser.empty()) {
                    SetWindowTextW(hStaticCurrentUser, L"👤 Текущий пользователь:");
                    if (hEditCurrentUser) SetWindowTextW(hEditCurrentUser, g_CurrentUser.c_str());
                }
                wstring k = LoadUserKey(g_CurrentUser);
                SetWindowTextW(hEditKey, k.c_str()); UpdateKeyLengthIndicator(hWnd);
                SetWindowTextW(hEditInput, L""); SetWindowTextW(hEditOutput, L"");
                SetWindowTextW(hBtnSaveStatus, L"💾 Сохранить");
                g_LastInputText.clear(); g_LastEncryptedFile.clear(); g_LastUser.clear();
                Exchange_UpdateAllRecipients(); Exchange_RefreshInboxList(); Exchange_RefreshOutboxList();
                break;
            }
            TranslateMessage(&m); DispatchMessage(&m);
        }
    }
}

// ------------------------
// Login window proc
// ------------------------
LRESULT CALLBACK LoginWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HWND cbUser, ePass, bOk, bNew, stTitle, stSub, bShow;
    static bool isNewUser = false;
    static HFONT fTitle = NULL, fSub = NULL;
    static vector<wstring> allUsers;

    auto FillUsers = [&]() {
        allUsers = ListSavedUsers();
        SendMessage(cbUser, CB_RESETCONTENT, 0, 0);
        for (auto& u : allUsers) SendMessage(cbUser, CB_ADDSTRING, 0, (LPARAM)u.c_str());
        };
    auto RebuildComboWithFilter = [&](const wstring& typed) {
        DWORD start = 0, end = 0; SendMessage(cbUser, CB_GETEDITSEL, (WPARAM)&start, (LPARAM)&end);
        SendMessage(cbUser, CB_RESETCONTENT, 0, 0);
        for (const auto& u : allUsers) {
            wstring H = u, N = typed; transform(H.begin(), H.end(), H.begin(), ::towlower); transform(N.begin(), N.end(), N.begin(), ::towlower);
            if (N.empty() || H.find(N) != wstring::npos) SendMessage(cbUser, CB_ADDSTRING, 0, (LPARAM)u.c_str());
        }
        SetWindowTextW(cbUser, typed.c_str());
        DWORD len = (DWORD)typed.size(); if (start > len) start = len; if (end > len) end = len;
        SendMessage(cbUser, CB_SETEDITSEL, 0, MAKELPARAM(start, end));
        SendMessage(cbUser, CB_SHOWDROPDOWN, TRUE, 0);
        };

    switch (msg) {
    case WM_CREATE: {
        isNewUser = false;
        HFONT base = CreateFontW(20, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        fTitle = CreateFontW(28, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        fSub = CreateFontW(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");

        const int W = 560, H = 320, M = 24, FW_ = W - 2 * M;
        RECT r; GetWindowRect(GetDesktopWindow(), &r);
        SetWindowPos(hWnd, nullptr, (r.right - W) / 2, (r.bottom - H) / 3, W, H, SWP_NOZORDER | SWP_NOACTIVATE);

        int y = M;
        stTitle = CreateWindowW(L"STATIC", L"Добро пожаловать", WS_CHILD | WS_VISIBLE | SS_CENTER, M, y, FW_, 34, hWnd, nullptr, hInst, nullptr);
        SendMessage(stTitle, WM_SETFONT, (WPARAM)fTitle, TRUE); y += 34 + 4;

        stSub = CreateWindowW(L"STATIC", L"Войдите или создайте нового пользователя", WS_CHILD | WS_VISIBLE | SS_CENTER, M, y, FW_, 20, hWnd, nullptr, hInst, nullptr);
        SendMessage(stSub, WM_SETFONT, (WPARAM)fSub, TRUE); y += 20 + 12;

        CreateWindowW(L"STATIC", L"👤 Имя пользователя", WS_CHILD | WS_VISIBLE | SS_LEFT, M, y, FW_, 18, hWnd, nullptr, hInst, nullptr); y += 18 + 6;

        cbUser = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | CBS_DROPDOWN | CBS_AUTOHSCROLL, M, y, FW_, 200, hWnd, (HMENU)ID_LOGIN_EDIT_USER, hInst, nullptr);
        ApplyExplorerTheme(cbUser);
        ApplyEditPadding(cbUser, hWnd);
        y += 32 + 12;

        CreateWindowW(L"STATIC", L"🔑 Пароль", WS_CHILD | WS_VISIBLE | SS_LEFT, M, y, FW_, 18, hWnd, nullptr, hInst, nullptr); y += 18 + 6;

        const int showW = 110;
        ePass = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_PASSWORD, M, y, FW_ - showW - 8, 30, hWnd, (HMENU)ID_LOGIN_EDIT_PASSWORD, hInst, nullptr);
        SendMessage(ePass, EM_SETPASSWORDCHAR, (WPARAM)L'•', 0);
        ApplyExplorerTheme(ePass);
        ApplyEditPadding(ePass, hWnd);

        bShow = CreateWindowW(L"BUTTON", L"Показать", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, M + (FW_ - showW), y + 5, showW, 20, hWnd, (HMENU)ID_LOGIN_SHOWPASS, hInst, nullptr);
        y += 30 + 16;

        const int BW = (FW_ - 10) / 2;
        bOk = CreateWindowW(L"BUTTON", L"🚪 Войти", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, M, y, BW, 32, hWnd, (HMENU)ID_LOGIN_OK, hInst, nullptr);
        bNew = CreateWindowW(L"BUTTON", L"📝 Создать", WS_CHILD | WS_VISIBLE, M + BW + 10, y, BW, 32, hWnd, (HMENU)ID_LOGIN_BTN_NEW_USER, hInst, nullptr);

        for (HWND c = GetWindow(hWnd, GW_CHILD); c; c = GetWindow(c, GW_HWNDNEXT)) SendMessage(c, WM_SETFONT, (WPARAM)base, TRUE);
        FillUsers(); SetFocus(cbUser); return 0;
    }
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hWnd, &ps);
        RECT rc; GetClientRect(hWnd, &rc);
        PaintGradientBackground(hdc, rc);
        RECT card = CombineRects(hWnd, { stTitle, stSub, cbUser, ePass, bOk, bNew, bShow }, ScaleByDpi(hWnd, 16));
        DrawSoftCard(hWnd, hdc, card);
        EndPaint(hWnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND: {
        RECT rc; GetClientRect(hWnd, &rc); PaintGradientBackground((HDC)wParam, rc); return 1;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam; SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT)); SetBkMode(hdc, TRANSPARENT); return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
    }
    case WM_CTLCOLOREDIT: {
        HDC hdc = (HDC)wParam; SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT)); SetBkColor(hdc, GetSysColor(COLOR_WINDOW)); return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
    }
    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wParam; SetTextColor(hdc, GetSysColor(COLOR_BTNTEXT)); SetBkColor(hdc, GetSysColor(COLOR_BTNFACE)); return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam), code = HIWORD(wParam);
        if (id == ID_LOGIN_EDIT_USER && code == CBN_EDITUPDATE) {
            wchar_t typed[256] = L""; GetWindowTextW(cbUser, typed, 256); RebuildComboWithFilter(typed); return 0;
        }
        if (id == ID_LOGIN_SHOWPASS) {
            const BOOL on = (SendMessage((HWND)lParam, BM_GETCHECK, 0, 0) == BST_CHECKED);
            SendMessage(ePass, EM_SETPASSWORDCHAR, (WPARAM)(on ? 0 : L'•'), 0); InvalidateRect(ePass, nullptr, TRUE); return 0;
        }
        if (id == ID_LOGIN_BTN_NEW_USER) {
            isNewUser = !isNewUser;
            SetWindowTextW(stTitle, isNewUser ? L"Создание пользователя" : L"Добро пожаловать");
            SetWindowTextW(bOk, isNewUser ? L"✅ Создать" : L"🚪 Войти");
            SetWindowTextW(bNew, isNewUser ? L"⬅ Назад" : L"📝 Создать");
            SetWindowTextW(cbUser, L""); SetWindowTextW(ePass, L""); SetFocus(cbUser); return 0;
        }
        if (id == ID_LOGIN_OK) {
            wchar_t ub[256], pb[128]; GetWindowTextW(cbUser, ub, 256); GetWindowTextW(ePass, pb, 128);
            wstring user = TrimWString(ub), pass = TrimWString(pb);
            if (user.empty()) { MessageBoxW(hWnd, L"Введите имя пользователя.", L"Ошибка", MB_OK | MB_ICONERROR); SetFocus(cbUser); return 0; }
            if (pass.empty()) { MessageBoxW(hWnd, L"Введите пароль.", L"Ошибка", MB_OK | MB_ICONERROR); SetFocus(ePass); return 0; }

            if (isNewUser) {
                if (UserExists(user)) { MessageBoxW(hWnd, L"Такой пользователь уже есть.", L"Ошибка", MB_OK | MB_ICONERROR); SetFocus(cbUser); return 0; }
                SaveUserPassword(user, pass); g_CurrentUser = user; PostMessage(nullptr, WM_USER + 1, 0, 0); DestroyWindow(hWnd);
            }
            else {
                if (UserExists(user) && VerifyPassword(user, pass)) { g_CurrentUser = user; PostMessage(nullptr, WM_USER + 1, 0, 0); DestroyWindow(hWnd); }
                else { MessageBoxW(hWnd, L"Неверные имя или пароль.", L"Ошибка", MB_OK | MB_ICONERROR); SetWindowTextW(ePass, L""); SetFocus(ePass); }
            }
            return 0;
        }
        break;
    }
    case WM_CLOSE: PostQuitMessage(0); return 0;
    case WM_DESTROY: if (fTitle) DeleteObject(fTitle); if (fSub) DeleteObject(fSub); return 0;
    }
    return DefWindowProc(hWnd, msg, wParam, lParam);
}

// --- мини-диалог для InputBox ---
struct IBState {
    HWND edit{};
    bool ok{ false };
    std::wstring result;
};

static LRESULT CALLBACK InputBoxWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    IBState* st = (IBState*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);
    switch (msg) {
    case WM_COMMAND: {
        switch (LOWORD(wParam)) {
        case IDOK: {
            if (st && st->edit) {
                wchar_t buf[512] = L"";
                GetWindowTextW(st->edit, buf, 512);
                st->result = TrimWString(buf);
                st->ok = true;
            }
            DestroyWindow(hWnd);
            return 0;
        }
        case IDCANCEL:
            if (st) st->ok = false;
            DestroyWindow(hWnd);
            return 0;
        }
        break;
    }
    case WM_CLOSE:
        if (st) st->ok = false;
        DestroyWindow(hWnd);
        return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ------------------------
// Keys Manager (ListView)
// ------------------------
LRESULT CALLBACK KeysWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    // ---- controls ----
    static HWND lv, bDel, bCopy, bRef, bSetCur, bRen, bExp, bImp, stUser;
    static HIMAGELIST hImg = NULL;

    // ---- DPI helper ----
    auto Dpi = [&](int px) -> int {
        UINT dpi = 96;
        if (HMODULE hU = GetModuleHandleW(L"user32")) {
            typedef UINT(WINAPI* FN)(HWND);
            static FN get = (FN)GetProcAddress(hU, "GetDpiForWindow");
            if (get) dpi = get(hWnd);
        }
        return MulDiv(px, (int)dpi, 96);
        };

    auto InputBox = [&](const std::wstring& title,
        const std::wstring& caption,
        const std::wstring& preset = L"") -> std::wstring
        {
            auto dpi = [&](int px)->int {
                UINT d = 96;
                if (HMODULE u = GetModuleHandleW(L"user32")) {
                    typedef UINT(WINAPI* FN)(HWND);
                    static FN get = (FN)GetProcAddress(u, "GetDpiForWindow");
                    if (get) d = get(hWnd);
                }
                return MulDiv(px, (int)d, 96);
                };

            const int M = dpi(12);
            const int BH = dpi(28);
            const int BW = dpi(92);
            const int GAP = dpi(10);
            const int CW = dpi(460);
            const int CH = dpi(150);

            RECT rcCli{ 0,0,CW,CH }, rcWnd{};
            AdjustWindowRectEx(&rcCli, WS_CAPTION | WS_SYSMENU | WS_POPUPWINDOW, FALSE, WS_EX_DLGMODALFRAME);
            rcWnd.right = rcCli.right - rcCli.left;
            rcWnd.bottom = rcCli.bottom - rcCli.top;

            HWND dlg = CreateWindowExW(
                WS_EX_DLGMODALFRAME, L"#32770", title.c_str(),
                WS_CAPTION | WS_SYSMENU | WS_POPUPWINDOW,
                CW_USEDEFAULT, CW_USEDEFAULT, rcWnd.right, rcWnd.bottom,
                hWnd, NULL, hInst, NULL);
            if (!dlg) return L"";

            // состояние + привяжем нашу WndProc
            auto* st = new IBState();
            SetWindowLongPtrW(dlg, GWLP_USERDATA, (LONG_PTR)st);
            SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)InputBoxWndProc);

            // центрирование
            RECT sr; GetWindowRect(GetDesktopWindow(), &sr);
            SetWindowPos(dlg, nullptr, (sr.right - rcWnd.right) / 2, (sr.bottom - rcWnd.bottom) / 3, 0, 0,
                SWP_NOSIZE | SWP_NOZORDER);

            // раскладка
            int y = M;
            CreateWindowW(L"STATIC", caption.c_str(), WS_CHILD | WS_VISIBLE,
                M, y, CW - 2 * M, dpi(18), dlg, 0, hInst, 0);
            y += dpi(18) + GAP;

            st->edit = CreateWindowW(L"EDIT", preset.c_str(),
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
                M, y, CW - 2 * M, dpi(28), dlg, (HMENU)1001, hInst, 0);
            ApplyExplorerTheme(st->edit);
            ApplyEditPadding(st->edit, dlg);
            y += dpi(28) + GAP;

            int yBtn = CH - M - BH;
            int xCancel = CW - M - BW;
            int xOK = xCancel - GAP - BW;
            HWND ok = CreateWindowW(L"BUTTON", L"OK",
                WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                xOK, yBtn, BW, BH, dlg, (HMENU)IDOK, hInst, 0);
            HWND ca = CreateWindowW(L"BUTTON", L"Отмена",
                WS_CHILD | WS_VISIBLE,
                xCancel, yBtn, BW, BH, dlg, (HMENU)IDCANCEL, hInst, 0);

            // шрифты
            SendMessage(st->edit, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(ok, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(ca, WM_SETFONT, (WPARAM)hFont, TRUE);

            ShowWindow(dlg, SW_SHOW);
            SetFocus(st->edit);

            // обычный модальный цикл
            MSG m;
            while (IsWindow(dlg) && GetMessage(&m, NULL, 0, 0)) {
                TranslateMessage(&m);
                DispatchMessage(&m);
            }

            std::wstring res;
            if (st->ok) res = st->result;
            delete st;
            return res;
        };

    // ---- list helpers ----
    auto FillList = [&]() {
        ListView_DeleteAllItems(lv);
        auto keys = LoadAllUserKeys(g_CurrentUser);
        int i = 0;
        for (auto& k : keys) {
            LVITEMW it{}; it.mask = LVIF_TEXT | LVIF_IMAGE; it.iItem = i;
            it.pszText = (LPWSTR)(k.when.empty() ? L"(без даты)" : k.when.c_str());
            it.iImage = k.isCurrent ? 1 : 0;
            int idx = ListView_InsertItem(lv, &it);
            ListView_SetItemText(lv, idx, 1, (LPWSTR)(k.label.empty() ? L"" : k.label.c_str()));
            ListView_SetItemText(lv, idx, 2, (LPWSTR)k.key.c_str());
            ListView_SetItemText(lv, idx, 3, (LPWSTR)(k.isCurrent ? L"Да" : L""));
            ++i;
        }
        };
    auto GetSelectedKey = [&]() -> wstring {
        int sel = ListView_GetNextItem(lv, -1, LVNI_SELECTED);
        if (sel == -1) return L"";
        wchar_t buf[256] = L""; ListView_GetItemText(lv, sel, 2, buf, 256);
        return NormalizeHex64(buf);
        };

    // ---- layout ----
    auto Layout = [&](HWND win) {
        RECT rc; GetClientRect(win, &rc);
        const int M = Dpi(10), rowGap = Dpi(8), btnH = Dpi(28);
        const int lvTop = M + Dpi(22);
        const int lvW = rc.right - 2 * M;
        const int lvH = max(LONG(Dpi(120)), (rc.bottom - lvTop) - (btnH + M + rowGap + Dpi(2)));

        SetWindowPos(stUser, NULL, M, M, lvW, Dpi(20), SWP_NOZORDER);
        SetWindowPos(lv, NULL, M, lvTop, lvW, (int)lvH, SWP_NOZORDER);

        // левая группа (копировать/удалить)
        int xL = M, y = rc.bottom - M - btnH;
        const int wCopy = Dpi(116), wDel = Dpi(100), gap = Dpi(8);

        SetWindowPos(bCopy, NULL, xL, y, wCopy, btnH, SWP_NOZORDER); xL += wCopy + gap;
        SetWindowPos(bDel, NULL, xL, y, wDel, btnH, SWP_NOZORDER);

        // правая группа
        HWND rightBtns[] = { bSetCur, bRen, bExp, bImp, bRef };
        int  widths[] = { Dpi(156), Dpi(136), Dpi(96), Dpi(96), Dpi(106) };
        int xR = rc.right - M;
        for (int i = 0; i < 5; ++i) {
            int w = widths[i]; xR -= w; SetWindowPos(rightBtns[i], NULL, xR, y, w, btnH, SWP_NOZORDER); xR -= gap;
        }

        // колонки
        int totalW = lvW;
        int wDate = Dpi(180), wLabel = Dpi(220), wCur = Dpi(80);
        int wKey = max(Dpi(260), totalW - (wDate + wLabel + wCur) - Dpi(20));
        ListView_SetColumnWidth(lv, 0, wDate);
        ListView_SetColumnWidth(lv, 1, wLabel);
        ListView_SetColumnWidth(lv, 3, wCur);
        ListView_SetColumnWidth(lv, 2, wKey);
        };

    switch (msg) {
    case WM_CREATE: {
        RECT sr; GetWindowRect(GetDesktopWindow(), &sr);
        int W = Dpi(900), H = Dpi(560);
        SetWindowPos(hWnd, nullptr, (sr.right - W) / 2, (sr.bottom - H) / 3, W, H, SWP_NOZORDER);

        InitCommonControls();

        stUser = CreateWindowW(L"STATIC",
            (L"Пользователь: " + g_CurrentUser).c_str(),
            WS_CHILD | WS_VISIBLE, 10, 10, 600, 20, hWnd, 0, hInst, 0);

        lv = CreateWindowW(WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL,
            10, 34, 860, 420, hWnd, (HMENU)ID_KEYS_LISTVIEW, hInst, 0);
        ListView_SetExtendedListViewStyle(lv, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);
        ApplyExplorerTheme(lv);

        hImg = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 2, 0);
        ImageList_AddIcon(hImg, LoadIcon(NULL, IDI_APPLICATION));
        ImageList_AddIcon(hImg, LoadIcon(NULL, IDI_ASTERISK));
        ListView_SetImageList(lv, hImg, LVSIL_SMALL);

        LVCOLUMNW col{ 0 }; col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        col.pszText = (LPWSTR)L"Дата";    col.cx = Dpi(180); col.iSubItem = 0; ListView_InsertColumn(lv, 0, &col);
        col.pszText = (LPWSTR)L"Метка";   col.cx = Dpi(220); col.iSubItem = 1; ListView_InsertColumn(lv, 1, &col);
        col.pszText = (LPWSTR)L"Ключ";    col.cx = Dpi(320); col.iSubItem = 2; ListView_InsertColumn(lv, 2, &col);
        col.pszText = (LPWSTR)L"Текущий"; col.cx = Dpi(80);  col.iSubItem = 3; ListView_InsertColumn(lv, 3, &col);

        // нижние кнопки
        bCopy = CreateWindowW(L"BUTTON", L"📋 Копировать", WS_CHILD | WS_VISIBLE, 0, 0, Dpi(116), Dpi(28), hWnd, (HMENU)ID_KEYS_COPY, hInst, 0);
        bDel = CreateWindowW(L"BUTTON", L"🗑 Удалить", WS_CHILD | WS_VISIBLE, 0, 0, Dpi(100), Dpi(28), hWnd, (HMENU)ID_KEYS_DELETE, hInst, 0);

        bRef = CreateWindowW(L"BUTTON", L"🔄 Обновить", WS_CHILD | WS_VISIBLE, 0, 0, Dpi(106), Dpi(28), hWnd, (HMENU)ID_KEYS_REFRESH, hInst, 0);
        bImp = CreateWindowW(L"BUTTON", L"📥 Импорт", WS_CHILD | WS_VISIBLE, 0, 0, Dpi(96), Dpi(28), hWnd, (HMENU)ID_KEYS_IMPORT, hInst, 0);
        bExp = CreateWindowW(L"BUTTON", L"📤 Экспорт", WS_CHILD | WS_VISIBLE, 0, 0, Dpi(96), Dpi(28), hWnd, (HMENU)ID_KEYS_EXPORT, hInst, 0);
        bRen = CreateWindowW(L"BUTTON", L"✏️ Переименовать", WS_CHILD | WS_VISIBLE, 0, 0, Dpi(136), Dpi(28), hWnd, (HMENU)ID_KEYS_RENAME, hInst, 0);
        bSetCur = CreateWindowW(L"BUTTON", L"⭐ Сделать текущим", WS_CHILD | WS_VISIBLE, 0, 0, Dpi(156), Dpi(28), hWnd, (HMENU)ID_KEYS_SET_CURRENT, hInst, 0);

        for (HWND c = GetWindow(hWnd, GW_CHILD); c; c = GetWindow(c, GW_HWNDNEXT))
            SendMessage(c, WM_SETFONT, (WPARAM)hFont, TRUE);

        FillList();
        Layout(hWnd);
        return 0;
    }

    case WM_SIZE:
        Layout(hWnd);
        return 0;

    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hWnd, &ps);
        RECT rc; GetClientRect(hWnd, &rc);
        PaintGradientBackground(hdc, rc);
        DrawSoftCard(hWnd, hdc, CombineRects(hWnd, { stUser, lv }, ScaleByDpi(hWnd, 14)));
        DrawSoftCard(hWnd, hdc, CombineRects(hWnd, { bCopy, bDel, bRef, bImp, bExp, bRen, bSetCur }, ScaleByDpi(hWnd, 12)));
        EndPaint(hWnd, &ps);
        return 0;
    }

    case WM_NOTIFY: {
        LPNMHDR hdr = (LPNMHDR)lParam;
        if (hdr->idFrom == ID_KEYS_LISTVIEW && hdr->code == NM_DBLCLK) {
            wstring key = GetSelectedKey();
            if (ValidateKey(key) && SetCurrentUserKey(g_CurrentUser, key)) {
                SetWindowTextW(hEditKey, key.c_str());
                UpdateKeyLengthIndicator(GetParent(hWnd));
                FillList(); Layout(hWnd);
            }
            return 0;
        }
        break;
    }

    case WM_COMMAND: {
        int id = LOWORD(wParam);
        switch (id) {
        case ID_KEYS_REFRESH:
            FillList(); Layout(hWnd); return 0;

        case ID_KEYS_SET_CURRENT: {
            wstring k = GetSelectedKey(); if (!ValidateKey(k)) return 0;
            if (SetCurrentUserKey(g_CurrentUser, k)) {
                SetWindowTextW(hEditKey, k.c_str());
                UpdateKeyLengthIndicator(GetParent(hWnd));
                FillList(); Layout(hWnd);
            }
            return 0;
        }

        case ID_KEYS_RENAME: {
            wstring k = GetSelectedKey(); if (!ValidateKey(k)) return 0;
            wstring lbl = InputBox(L"Переименовать ключ", L"Введите новую метку:", L"");
            if (!lbl.empty()) { UpdateUserKeyLabel(g_CurrentUser, k, lbl); FillList(); Layout(hWnd); }
            return 0;
        }

        case ID_KEYS_COPY: {
            wstring k = GetSelectedKey(); if (!ValidateKey(k)) return 0;
            CopyKeyToClipboard(hWnd, k);
            return 0;
        }

        case ID_KEYS_EXPORT: {
            wstring k = GetSelectedKey(); if (!ValidateKey(k)) return 0;
            if (!ExportUserKeyToFile(hWnd, k))
                MessageBoxW(hWnd, L"Не удалось экспортировать ключ.", L"Ошибка", MB_OK | MB_ICONERROR);
            return 0;
        }

        case ID_KEYS_IMPORT: {
            if (!ImportUserKeyFromFile(hWnd, g_CurrentUser)) {
                MessageBoxW(hWnd, L"Не удалось импортировать ключ.", L"Ошибка", MB_OK | MB_ICONERROR);
            }
            else {
                FillList(); Layout(hWnd);
            }
            return 0;
        }

        case ID_KEYS_DELETE: {
            wstring k = GetSelectedKey(); if (!ValidateKey(k)) return 0;
            if (MessageBoxW(hWnd, L"Удалить выбранный ключ?", L"Подтверждение", MB_YESNO | MB_ICONQUESTION) == IDYES) {
                if (!DeleteUserKey(g_CurrentUser, k))
                    MessageBoxW(hWnd, L"Не удалось удалить ключ.", L"Ошибка", MB_OK | MB_ICONERROR);
                FillList(); Layout(hWnd);
            }
            return 0;
        }
        }
        break;
    }

    case WM_DESTROY:
        if (hImg) { ImageList_Destroy(hImg); hImg = NULL; }
        return 0;

    case WM_ERASEBKGND: {
        RECT rc; GetClientRect(hWnd, &rc);
        PaintGradientBackground((HDC)wParam, rc);
        return 1;
    }
    }

    return DefWindowProc(hWnd, msg, wParam, lParam);
}

ATOM RegisterKeysClass(HINSTANCE hInstance) {
    WNDCLASSEXW w = { sizeof(WNDCLASSEXW) };
    w.style = CS_HREDRAW | CS_VREDRAW;
    w.lpfnWndProc = KeysWndProc;
    w.hInstance = hInstance;
    w.hCursor = LoadCursor(NULL, IDC_ARROW);
    w.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    w.lpszClassName = L"KEYS_GUI";
    w.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    w.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
    return RegisterClassExW(&w);
}

void OpenKeysManager(HWND parent) {
    if (g_CurrentUser.empty()) {
        MessageBoxW(parent, L"Сначала войдите в систему.", L"Ошибка", MB_OK | MB_ICONERROR);
        return;
    }
    HWND dlg = CreateWindowW(L"KEYS_GUI", L"Ключи пользователя",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 900, 560,
        parent, NULL, hInst, NULL);
    if (dlg) { ShowWindow(dlg, SW_SHOW); UpdateWindow(dlg); }
}

// ------------------------
// Exchange window
// ------------------------
LRESULT CALLBACK ExchangeWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        const int M = 12;
        const int H = 28;
        int y = M;

        CreateWindowW(L"STATIC", L"🤝 Обмен файлами",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M, y, 260, H, hWnd, (HMENU)-1, hInst, NULL);
        y += H + 8;

        CreateWindowW(L"STATIC", L"Получатель:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M, y, 90, H, hWnd, (HMENU)-1, hInst, NULL);

        hExCombo = CreateWindowW(L"COMBOBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | CBS_DROPDOWN | CBS_AUTOHSCROLL,
            M + 90 + 6, y, 220, 200,
            hWnd, (HMENU)ID_EXCH_COMBO, hInst, NULL);
        ApplyExplorerTheme(hExCombo);
        ApplyEditPadding(hExCombo, hWnd);

        hExSend = CreateWindowW(L"BUTTON", L"📨 Отправить",
            WS_CHILD | WS_VISIBLE,
            M + 90 + 6 + 220 + 8, y, 140, H,
            hWnd, (HMENU)ID_EXCH_SEND, hInst, NULL);
        y += H + 10;

        CreateWindowW(L"STATIC", L"📬 Входящие:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M, y, 140, H, hWnd, (HMENU)-1, hInst, NULL);

        hExRefresh = CreateWindowW(L"BUTTON", L"🔄 Обновить",
            WS_CHILD | WS_VISIBLE,
            M + 140 + 8, y, 120, H,
            hWnd, (HMENU)ID_EXCH_REFRESH, hInst, NULL);

        // узнаём ширину окна, чтобы прижать кнопку к правому краю
        RECT rc;
        GetClientRect(hWnd, &rc);
        int halfW = (rc.right - 3 * M) / 2;

        // кнопка "Папка обмена" справа, на одной строке с "Входящие"
        int folderWidth = 140;
        hExOpenFolder = CreateWindowW(L"BUTTON", L"📂 Папка обмена",
            WS_CHILD | WS_VISIBLE,
            rc.right - M - folderWidth, y, folderWidth, H,
            hWnd, (HMENU)ID_EXCH_OPEN_FOLDER, hInst, NULL);

        // списки начинаются строкой ниже
        int listTop = y + H + 6;
        int listH = rc.bottom - listTop - M;

        // левый список — входящие
        hExInList = CreateWindowW(L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY,
            M, listTop, halfW, listH,
            hWnd, (HMENU)ID_EXCH_IN_LIST, hInst, NULL);
        ApplyExplorerTheme(hExInList);

        // заголовок "Исходящие:" над правым списком
        CreateWindowW(L"STATIC", L"📤 Исходящие:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M + halfW + M, y, 140, H, hWnd, (HMENU)-1, hInst, NULL);

        // правый список — исходящие
        hExOutList = CreateWindowW(L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY,
            M + halfW + M, listTop, halfW, listH,
            hWnd, (HMENU)ID_EXCH_OUT_LIST, hInst, NULL);
        ApplyExplorerTheme(hExOutList);


        // шрифты
        for (HWND c = GetWindow(hWnd, GW_CHILD); c; c = GetWindow(c, GW_HWNDNEXT))
            SendMessage(c, WM_SETFONT, (WPARAM)hFont, TRUE);

        Exchange_UpdateAllRecipients();
        Exchange_RefreshInboxList();
        Exchange_RefreshOutboxList();
        return 0;
    }
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hWnd, &ps);
        RECT rc; GetClientRect(hWnd, &rc);
        PaintGradientBackground(hdc, rc);
        RECT lists = CombineRects(hWnd, { hExCombo, hExSend, hExInList, hExOutList, hExRefresh, hExOpenFolder }, ScaleByDpi(hWnd, 14));
        DrawSoftCard(hWnd, hdc, lists);
        EndPaint(hWnd, &ps);
        return 0;
    }

    case WM_ERASEBKGND: {
        RECT rc; GetClientRect(hWnd, &rc);
        PaintGradientBackground((HDC)wParam, rc);
        return 1;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
    }
    case WM_CTLCOLOREDIT: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));
        SetBkColor(hdc, GetSysColor(COLOR_WINDOW));
        return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
    }
    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, GetSysColor(COLOR_BTNTEXT));
        SetBkColor(hdc, GetSysColor(COLOR_BTNFACE));
        return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
    }

    case WM_DRAWITEM:
        if (DrawStyledButton((LPDRAWITEMSTRUCT)lParam)) return TRUE;
        break;

    case WM_COMMAND: {
        int id = LOWORD(wParam);
        int ev = HIWORD(wParam);

        if (id == ID_EXCH_COMBO && ev == CBN_EDITUPDATE) {
            wchar_t typed[256] = L"";
            GetWindowTextW(hExCombo, typed, 256);
            Exchange_RebuildRecipientsCombo(typed, true);
            return 0;
        }

        if ((id == ID_EXCH_IN_LIST || id == ID_EXCH_OUT_LIST) && ev == LBN_DBLCLK) {
            int sel = (int)SendMessage((HWND)lParam, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR) {
                wchar_t name[512];
                SendMessage((HWND)lParam, LB_GETTEXT, sel, (LPARAM)name);
                wstring base = GetExchangeDir() + L"\\" + g_CurrentUser +
                    (id == ID_EXCH_IN_LIST ? L"\\Inbox\\" : L"\\Outbox\\");
                wstring path = base + name;
                ifstream in(path, ios::binary);
                if (in) {
                    string buf((istreambuf_iterator<char>(in)), {});
                    wstring fileData = UTF8ToWString(buf);
                    wstring keyFromFile, payload;
                    if (Exchange_ExtractKeyAndPayload(fileData, keyFromFile, payload)) {
                        SetWindowTextW(hEditKey, keyFromFile.c_str());
                        UpdateKeyLengthIndicator(GetParent(hWnd));
                        SetWindowTextW(hEditInput, payload.c_str());
                    }
                    else {
                        SetWindowTextW(hEditInput, fileData.c_str());
                    }
                    MessageBoxW(hWnd, L"Текст загружен в левое поле основного окна.", L"Обмен", MB_OK | MB_ICONINFORMATION);
                }
            }
            return 0;
        }

        switch (id) {
        case ID_EXCH_REFRESH:
            Exchange_UpdateAllRecipients();
            Exchange_RefreshInboxList();
            Exchange_RefreshOutboxList();
            return 0;

        case ID_EXCH_OPEN_FOLDER: {
            if (g_CurrentUser.empty()) {
                MessageBoxW(hWnd, L"Сначала войдите в систему.", L"Ошибка", MB_OK | MB_ICONERROR);
                break;
            }
            wstring dir = GetExchangeDir() + L"\\" + g_CurrentUser;
            if (!OpenFolderInExplorer(dir))
                MessageBoxW(hWnd, L"Не удалось открыть папку обмена.", L"Ошибка", MB_OK | MB_ICONERROR);
            return 0;
        }

        case ID_EXCH_SEND: {
            if (g_CurrentUser.empty()) {
                MessageBoxW(hWnd, L"Сначала войдите в систему.", L"Ошибка", MB_OK | MB_ICONERROR);
                break;
            }

            int sel = (int)SendMessage(hExCombo, CB_GETCURSEL, 0, 0);
            wchar_t recBuf[256] = L"";
            if (sel != CB_ERR)
                SendMessage(hExCombo, CB_GETLBTEXT, sel, (LPARAM)recBuf);
            else
                GetWindowTextW(hExCombo, recBuf, 256);

            wstring recipient = TrimWString(recBuf);
            if (recipient.empty()) {
                MessageBoxW(hWnd, L"Выберите получателя.", L"Ошибка", MB_OK | MB_ICONERROR);
                break;
            }

            wchar_t inputBuf[20000];
            GetWindowTextW(hEditInput, inputBuf, 20000);
            wstring plain = TrimWString(inputBuf);
            if (plain.empty()) {
                MessageBoxW(hWnd, L"Нет текста для отправки (левое поле основного окна пустое).", L"Ошибка", MB_OK | MB_ICONERROR);
                break;
            }

            int mode = (int)SendMessage(hComboMode, CB_GETCURSEL, 0, 0);
            if (Exchange_SendTo(recipient, plain, mode)) {
                MessageBoxW(hWnd, L"Отправлено!", L"Обмен", MB_OK | MB_ICONINFORMATION);
                Exchange_RefreshOutboxList();
            }
            else {
                MessageBoxW(hWnd, L"Не удалось отправить (нет ключа у получателя или ошибка записи).", L"Обмен", MB_OK | MB_ICONERROR);
            }
            return 0;
        }
        }
        break;
    }

    case WM_DESTROY:
        // обнуляем глобальные HWND обмена
        hExCombo = hExSend = hExRefresh = hExInList = hExOutList = hExOpenFolder = NULL;
        return 0;
    }

    return DefWindowProc(hWnd, msg, wParam, lParam);
}

ATOM RegisterExchangeClass(HINSTANCE hInstance) {
    WNDCLASSEXW w = { sizeof(WNDCLASSEXW) };
    w.style = CS_HREDRAW | CS_VREDRAW;
    w.lpfnWndProc = ExchangeWndProc;
    w.hInstance = hInstance;
    w.hCursor = LoadCursor(NULL, IDC_ARROW);
    w.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    w.lpszClassName = L"EXCHANGE_GUI";
    w.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    w.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
    return RegisterClassExW(&w);
}

void OpenExchangeWindow(HWND parent) {
    if (g_CurrentUser.empty()) {
        MessageBoxW(parent, L"Сначала войдите в систему.", L"Ошибка", MB_OK | MB_ICONERROR);
        return;
    }
    HWND dlg = CreateWindowW(L"EXCHANGE_GUI", L"Обмен файлами",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 700, 400,
        parent, NULL, hInst, NULL);
    if (dlg) { ShowWindow(dlg, SW_SHOW); UpdateWindow(dlg); }
}

// ------------------------
// Main window proc
// ------------------------
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        // Меню
        HMENU hMenuBar = CreateMenu();
        HMENU hFileMenu = CreateMenu();
        HMENU hToolsMenu = CreateMenu();

        AppendMenuW(hFileMenu, MF_STRING, IDM_FILE_OPEN, L"Открыть...");
        AppendMenuW(hFileMenu, MF_STRING, IDM_FILE_SAVE, L"Сохранить");
        AppendMenuW(hFileMenu, MF_SEPARATOR, 0, NULL);
        AppendMenuW(hFileMenu, MF_STRING, IDM_FILE_EXIT, L"Выход");

        AppendMenuW(hToolsMenu, MF_STRING, ID_BTN_OPEN_KEYS_FILE, L"Менеджер ключей");
        AppendMenuW(hToolsMenu, MF_STRING, IDM_TOOLS_EXCHANGE, L"Обмен файлами");

        AppendMenuW(hMenuBar, MF_POPUP, (UINT_PTR)hFileMenu, L"&Файл");
        AppendMenuW(hMenuBar, MF_POPUP, (UINT_PTR)hToolsMenu, L"&Сервис");

        SetMenu(hWnd, hMenuBar);

        // Шрифты
        hFont = CreateFontW(18, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
            DEFAULT_PITCH, L"Segoe UI");
        hFontTitle = CreateFontW(28, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
            DEFAULT_PITCH, L"Segoe UI");
        const int W = 1200;
        const int M = UI_PAD;
        const int HALF = (W - M * 3) / 2;
        int y = M;

        HWND title = CreateWindowW(L"STATIC", L"ГОСТ 28147-89 🔒",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            0, y, W, 36, hWnd, (HMENU)-1, hInst, NULL);
        SendMessage(title, WM_SETFONT, (WPARAM)hFontTitle, TRUE);
        y += 36;

        y += UI_GAP;

        hStaticCurrentUser = CreateWindowW(L"STATIC", L"👤 Текущий пользователь:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M, y, 200, UI_H, hWnd, (HMENU)ID_STATIC_CURRENT_USER, hInst, NULL);

        hEditCurrentUser = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER |
            ES_READONLY | ES_AUTOHSCROLL,
            M + 200 + 8, y, 220, UI_H,
            hWnd, (HMENU)ID_EDIT_CURRENT_USER, hInst, NULL);
        ApplyExplorerTheme(hEditCurrentUser);
        ApplyEditPadding(hEditCurrentUser, hWnd);

        CreateStyledButton(hWnd, ID_BTN_CHANGE_USER, L"🔄 Сменить пользователя",
            M + 200 + 8 + 220 + 8, y, UI_BTN + 30, UI_H, false);
        y += UI_H + UI_GAP;

        CreateWindowW(L"STATIC", L"🔑 Ключ (64 HEX):",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M, y, 140, UI_H, hWnd, (HMENU)-1, hInst, NULL);

        hEditKey = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            M + 140 + 8, y, 600, UI_H,
            hWnd, (HMENU)ID_EDIT_KEY, hInst, NULL);
        ApplyExplorerTheme(hEditKey);
        ApplyEditPadding(hEditKey, hWnd);

        hStaticKeyLength = CreateWindowW(L"STATIC", L"0/64 символов",
            WS_CHILD | WS_VISIBLE,
            M + 140 + 8 + 600 + 8, y,
            120, UI_H, hWnd, (HMENU)ID_STATIC_KEY_LENGTH, hInst, NULL);

        CreateStyledButton(hWnd, ID_BTN_GENERATE_KEY, L"🔧 Сгенерировать",
            M + 140 + 8 + 600 + 8 + 120 + 8,
            y,
            120, UI_H, false);

        RECT rc;
        GetClientRect(hWnd, &rc);

        // Высота = доступная высота минус отступы и блок кнопок снизу
        int BOX_H = rc.bottom - y - UI_H * 5 - 40;
        if (BOX_H < 200) BOX_H = 200;

        CreateWindowW(L"STATIC", L"📄 Исходный текст:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M, y, 140, UI_H, hWnd, (HMENU)-1, hInst, NULL);

        hEditInput = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER |
            ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
            M, y + UI_H + 6, HALF - M, BOX_H,
            hWnd, (HMENU)ID_EDIT_INPUT, hInst, NULL);
        ApplyExplorerTheme(hEditInput);
        ApplyEditPadding(hEditInput, hWnd);

        CreateWindowW(L"STATIC", L"📤 Результат:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M + HALF + M, y, 100, UI_H, hWnd, (HMENU)-1, hInst, NULL);

        hEditOutput = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER |
            ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
            M + HALF + M, y + UI_H + 6, HALF - M, BOX_H,
            hWnd, (HMENU)ID_EDIT_OUTPUT, hInst, NULL);
        ApplyExplorerTheme(hEditOutput);
        ApplyEditPadding(hEditOutput, hWnd);
        y += BOX_H + UI_H + 12;

        // Блок режимов (обмен вынесен в отдельное окно через меню)
        CreateWindowW(L"STATIC", L"⚙️ Режим:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            M, y, 80, UI_H, hWnd, (HMENU)-1, hInst, NULL);

        hComboMode = CreateWindowW(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
            M + 80 + 8, y, 220, 200,
            hWnd, (HMENU)200, hInst, NULL);
        ApplyExplorerTheme(hComboMode);

        SendMessage(hComboMode, CB_ADDSTRING, 0, (LPARAM)L"ECB (простой)");
        SendMessage(hComboMode, CB_ADDSTRING, 0, (LPARAM)L"CBC (безопасный)");
        SendMessage(hComboMode, CB_ADDSTRING, 0, (LPARAM)L"CFB (потоковый)");
        SendMessage(hComboMode, CB_SETCURSEL, 1, 0);

        hBtnHelp = CreateStyledButton(hWnd, ID_BTN_HELP, L"❓ Справка",
            M + 80 + 8 + 220 + 8, y, 100, UI_H, false);
        y += UI_H + 8;

        int utilityX = M;
        int utilityGap = 14;
        hBtnCopy = CreateStyledButton(hWnd, ID_BTN_COPY, L"📋 Копировать ключ",
            utilityX, y, 200, UI_H, false);
        utilityX += 200 + utilityGap;

        hBtnSaveStatus = CreateStyledButton(hWnd, ID_BTN_SAVE_STATUS, L"💾 Сохранить",
            utilityX, y, 160, UI_H, false);
        utilityX += 160 + utilityGap;

        hBtnOpenOutputDir = CreateStyledButton(hWnd, ID_BTN_OPEN_OUTPUT_DIR, L"📂 Выходные файлы",
            utilityX, y, 200, UI_H, false);

        y += UI_H + 8;

        int WBTN = 170;
        int GAP = 14;
        int total = 4 * WBTN + 3 * GAP;
        int X = (W - total) / 2;

        CreateStyledButton(hWnd, ID_BTN_ENCRYPT, L"🔒 Зашифровать",
            X, y, WBTN, UI_H, true);

        CreateStyledButton(hWnd, ID_BTN_DECRYPT_LAST, L"🔓 Расшифровать последний",
            X + WBTN + GAP, y, WBTN + 40, UI_H,
            true);

        CreateStyledButton(hWnd, ID_BTN_DECRYPT, L"🔓 Расшифровать",
            X + WBTN + GAP + (WBTN + 40) + GAP, y,
            WBTN, UI_H, true);

        CreateStyledButton(hWnd, ID_BTN_CLEAR, L"✂️ Очистить",
            X + WBTN + GAP + (WBTN + 40) + GAP + WBTN + GAP, y,
            WBTN, UI_H, false);

        HWND c = GetWindow(hWnd, GW_CHILD);
        while (c) {
            SendMessage(c, WM_SETFONT, (WPARAM)hFont, TRUE);
            c = GetWindow(c, GW_HWNDNEXT);
        }

        InitToolTips(hWnd);

        // обновим список пользователей для обмена (окно обменобмена использует эти данные)
        Exchange_UpdateAllRecipients();
        Exchange_RefreshInboxList();
        Exchange_RefreshOutboxList();
        break;
    }

    case WM_USER + 1: {
        if (!g_CurrentUser.empty()) {
            if (hStaticCurrentUser)
                SetWindowTextW(hStaticCurrentUser, L"👤 Текущий пользователь:");
            if (hEditCurrentUser)
                SetWindowTextW(hEditCurrentUser, g_CurrentUser.c_str());

            wstring k = LoadUserKey(g_CurrentUser);
            SetWindowTextW(hEditKey, k.c_str());
            UpdateKeyLengthIndicator(hWnd);

            SetWindowTextW(hEditInput, L"");
            SetWindowTextW(hEditOutput, L"");
            SetWindowTextW(hBtnSaveStatus, L"💾 Сохранить");
            g_LastInputText.clear();
            g_LastEncryptedFile.clear();
            g_LastUser.clear();

            Exchange_UpdateAllRecipients();
            Exchange_RefreshInboxList();
            Exchange_RefreshOutboxList();
        }
        break;
    }

    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hWnd, &ps);
        RECT rc; GetClientRect(hWnd, &rc);
        PaintGradientBackground(hdc, rc);

        RECT topCard = CombineRects(hWnd,
            { hStaticCurrentUser, hEditCurrentUser, hEditKey, hStaticKeyLength }, ScaleByDpi(hWnd, 16));
        DrawSoftCard(hWnd, hdc, topCard);

        RECT textCard = CombineRects(hWnd, { hEditInput, hEditOutput }, ScaleByDpi(hWnd, 18));
        DrawSoftCard(hWnd, hdc, textCard);

        RECT actions = CombineRects(hWnd, { hComboMode, hBtnHelp, hBtnSaveStatus, hBtnOpenOutputDir, hBtnCopy }, ScaleByDpi(hWnd, 14));
        DrawSoftCard(hWnd, hdc, actions);

        EndPaint(hWnd, &ps);
        return 0;
    }

    case WM_ERASEBKGND: {
        RECT rc; GetClientRect(hWnd, &rc);
        PaintGradientBackground((HDC)wParam, rc);
        return 1;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
    }
    case WM_CTLCOLOREDIT: {
        HDC  hdc = (HDC)wParam;
        HWND he = (HWND)lParam;
        SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));

        if (he == hEditKey) {
            return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
        }

        SetBkColor(hdc, GetSysColor(COLOR_WINDOW));
        return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
    }
    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, GetSysColor(COLOR_BTNTEXT));
        SetBkColor(hdc, GetSysColor(COLOR_BTNFACE));
        return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
    }

    case WM_DRAWITEM:
        if (DrawStyledButton((LPDRAWITEMSTRUCT)lParam)) return TRUE;
        break;

    case WM_COMMAND: {
        int id = LOWORD(wParam);
        int ev = HIWORD(wParam);

        // смена режима
        if (id == 200 && ev == CBN_SELCHANGE) {
            int sel = (int)SendMessage(hComboMode, CB_GETCURSEL, 0, 0);
            return 0;
        }

        if (id == ID_EDIT_KEY && (ev == EN_CHANGE || ev == EN_KILLFOCUS)) {
            if (ev == EN_KILLFOCUS) {
                wchar_t b[256];
                GetWindowTextW(hEditKey, b, 256);
                wstring n = NormalizeHex64(b);
                SetWindowTextW(hEditKey, n.c_str());
                if (!g_CurrentUser.empty() && ValidateKey(n))
                    SaveUserKey(g_CurrentUser, n, L"");
            }
            UpdateKeyLengthIndicator(hWnd);
            return 0;
        }

        switch (id) {
            // Меню Файл
        case IDM_FILE_OPEN:
            HandleOpenFile(hWnd);
            return 0;
        case IDM_FILE_SAVE:
            HandleSaveResult(hWnd);
            return 0;
        case IDM_FILE_EXIT:
            SendMessage(hWnd, WM_CLOSE, 0, 0);
            return 0;

            // Меню Сервис
        case ID_BTN_OPEN_KEYS_FILE:
            OpenKeysManager(hWnd);
            return 0;
        case IDM_TOOLS_EXCHANGE:
            OpenExchangeWindow(hWnd);
            return 0;

        case ID_BTN_CHANGE_USER:
            ChangeUser(hWnd);
            break;

        case ID_BTN_GENERATE_KEY: {
            wstring k = GenerateRandomKey();
            SetWindowTextW(hEditKey, k.c_str());
            UpdateKeyLengthIndicator(hWnd);
            if (!g_CurrentUser.empty())
                SaveUserKey(g_CurrentUser, k, L"");
            break;
        }

        case ID_BTN_COPY: {
            wchar_t b[256];
            GetWindowTextW(hEditKey, b, 256);
            CopyKeyToClipboard(hWnd, NormalizeHex64(b));
            break;
        }

        case ID_BTN_CLEAR: {
            SetWindowTextW(hEditInput, L"");
            SetWindowTextW(hEditOutput, L"");
            SetWindowTextW(hBtnSaveStatus, L"💾 Сохранить");
            g_LastInputText = L"";
            break;
        }

        case ID_BTN_ENCRYPT:
            HandleEncryptDecrypt(hWnd, true, false);
            break;

        case ID_BTN_DECRYPT:
            HandleEncryptDecrypt(hWnd, false, false);
            break;

        case ID_BTN_DECRYPT_LAST:
            HandleEncryptDecrypt(hWnd, false, true);
            break;

        case ID_BTN_SAVE_STATUS:
            HandleSaveResult(hWnd);
            break;

        case ID_BTN_OPEN_OUTPUT_DIR: {
            wstring target = GetOutputDir();
            if (!g_CurrentUser.empty())
                target += L"\\" + g_CurrentUser;

            if (!OpenFolderInExplorer(target))
                MessageBoxW(hWnd, L"Не удалось открыть папку.", L"Ошибка", MB_OK | MB_ICONERROR);
            break;
        }

        case ID_BTN_HELP: {
            wstring help =
                L"Справка по использованию программы\n\n"
                L"1. Текущий пользователь\n"
                L"   • В верхней части окна указано имя активного пользователя.\n"
                L"   • Чтобы войти под другим пользователем — нажмите «🔄 Сменить пользователя».\n"
                L"   • Для каждого пользователя хранятся свои ключи и пароли.\n\n"
                L"2. Ключи шифрования\n"
                L"   • Менеджер ключей (через меню «Сервис → Менеджер ключей»)\n"
                L"     позволяет хранить несколько ключей с метками и помечать ⭐ текущий.\n\n"
                L"3. Шифрование/расшифрование\n"
                L"   • Введите текст слева или загрузите файл.\n"
                L"   • Выберите режим (рекомендуется CBC).\n"
                L"   • Нажмите «🔒 Зашифровать» или «🔓 Расшифровать».\n"
                L"   • Результат автосохраняется в «Выходные файлы\\\\<пользователь>».\n\n"
                L"4. Обмен\n"
                L"   • Открывается через меню «Сервис → Обмен файлами».\n"
                L"   • Использует левое поле основного окна как текст для отправки.\n\n"
                L"Примечание: Контейнер GOST0 содержит режим и IV; режим выбирать при расшифровании не требуется.";
            MessageBoxW(hWnd, help.c_str(), L"Справка", MB_OK | MB_ICONINFORMATION);
            break;
        }
        }
        break;
    }

    case WM_DESTROY:
        DestroyWindow(hTooltip);
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

// ------------------------
// Register classes
// ------------------------
ATOM MyRegisterClass(HINSTANCE hInstance) {
    WNDCLASSEXW w = { sizeof(WNDCLASSEXW) };
    w.style = CS_HREDRAW | CS_VREDRAW;
    w.lpfnWndProc = WndProc;
    w.hInstance = hInstance;
    w.hCursor = LoadCursor(NULL, IDC_ARROW);
    w.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    w.lpszClassName = L"GOST_GUI";
    w.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    w.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
    return RegisterClassExW(&w);
}

ATOM RegisterLoginClass(HINSTANCE hInstance) {
    WNDCLASSEXW w = { sizeof(WNDCLASSEXW) };
    w.style = CS_HREDRAW | CS_VREDRAW;
    w.lpfnWndProc = LoginWndProc;
    w.hInstance = hInstance;
    w.hCursor = LoadCursor(NULL, IDC_ARROW);
    w.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    w.lpszClassName = L"LOGIN_GUI";
    w.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    w.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
    return RegisterClassExW(&w);
}

// ------------------------
// Create main window after login
// ------------------------
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow) {
    HWND hWnd = CreateWindowExW(WS_EX_COMPOSITED,
        L"GOST_GUI", L"ГОСТ 28147-89",
        WS_OVERLAPPEDWINDOW ^ WS_THICKFRAME ^ WS_MAXIMIZEBOX,
        CW_USEDEFAULT, 0, 1200, 800,
        NULL, NULL, hInstance, NULL);
    if (!hWnd) return FALSE;
    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);
    return TRUE;
}

// ------------------------
// Entry point
// ------------------------
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    hInst = hInstance;

    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_WIN95_CLASSES | ICC_STANDARD_CLASSES };
    if (!InitCommonControlsEx(&icc)) return FALSE;

    BASE_DIR_PATH = GetExePath();
    EnsureDirectoryExists(BASE_DIR_PATH);
    EnsureDirectoryExists(GetOutputDir());
    EnsureDirectoryExists(GetExchangeDir());

    if (!RegisterLoginClass(hInstance))   return FALSE;
    if (!MyRegisterClass(hInstance))      return FALSE;
    if (!RegisterKeysClass(hInstance))    return FALSE;
    if (!RegisterExchangeClass(hInstance))return FALSE;

    HWND login = CreateWindowW(L"LOGIN_GUI", L"Вход в программу",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 560, 380,
        NULL, NULL, hInstance, NULL);
    if (!login) return FALSE;
    ShowWindow(login, SW_SHOW);
    UpdateWindow(login);

    HWND mainWnd = NULL;
    MSG msg;

    while (GetMessage(&msg, NULL, 0, 0)) {
        if (msg.message == WM_USER + 1) {
            if (!mainWnd) {
                if (!InitInstance(hInstance, nCmdShow)) return FALSE;
                mainWnd = FindWindowW(L"GOST_GUI", NULL);
                if (!mainWnd) return FALSE;
            }
            SendMessage(mainWnd, WM_USER + 1, 0, 0);
            continue;
        }
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    if (hFont)        DeleteObject(hFont);
    if (hFontTitle)   DeleteObject(hFontTitle);

    return (int)msg.wParam;
}
