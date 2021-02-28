#include <sdkddkver.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>
#include <windowsx.h>
#include <assert.h>
#include <strsafe.h>
#include <stdlib.h>
#include <Uxtheme.h>
#pragma comment(lib, "uxtheme.lib")

#include "Win32Toolbox.h"


#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")


// Forward declarations of functions included in this code module:
static ATOM                MyRegisterClass(HINSTANCE hInstance);
static LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
static void                UpdateCapturedContent(HWND hWnd);

static HINSTANCE hInst;


int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                      _In_opt_ HINSTANCE hPrevInstance,
                      _In_ LPWSTR    lpCmdLine,
                      _In_ int       nCmdShow)
{
	hInst = hInstance;

	// Initialize global strings
	ATOM Atom_MainWindow = MyRegisterClass(hInstance);

	HWND hWnd = CreateWindowExW(0, MAKEINTATOM(Atom_MainWindow), L"Clipboard Monitor", WS_OVERLAPPEDWINDOW | WS_VSCROLL | WS_HSCROLL,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);
	assert(hWnd != nullptr);
	ShowWindow(hWnd, nCmdShow);

	MSG msg;
	// Main message loop:
	while (GetMessageW(&msg, nullptr, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessageW(&msg);
	}

	return (int)msg.wParam;
}


ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEXW wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style          = CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc    = WndProc;
	wcex.cbClsExtra     = 0;
	wcex.cbWndExtra     = 0;
	wcex.hInstance      = hInstance;
	wcex.hIcon          = nullptr;
	wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
	wcex.hbrBackground  = nullptr;
	wcex.lpszMenuName   = nullptr;
	wcex.lpszClassName  = L"ClipboardMonitorMainWindow";
	wcex.hIconSm        = nullptr;

	return RegisterClassExW(&wcex);
}


#define IDM_CLEAR_CLIPBOARD 100
#define IDM_REFRESH 101
#define IDM_TOGGLE_AUTO 102


static HBITMAP CurrentImage;
static LONG CurrentImageWidth;
static LONG CurrentImageHeight;

static LPWSTR CurrentText;
static HWND CurrentEditControl;


static HBITMAP CreateCopyBitmap(HDC hdc, LONG Width, LONG Height)
{
	BITMAPINFOHEADER BitmapInfo = {};
	BitmapInfo.biSize = sizeof(BitmapInfo);
	BitmapInfo.biWidth = Width;
	BitmapInfo.biHeight = Height;
	BitmapInfo.biPlanes = 1;
	BitmapInfo.biBitCount = 24;
	BitmapInfo.biCompression = BI_RGB;
	void *Bits = nullptr;
	HBITMAP Bitmap = CreateDIBSection(hdc, (BITMAPINFO *)&BitmapInfo, DIB_RGB_COLORS, &Bits, nullptr, 0);
	assert(Bitmap != nullptr);
	return Bitmap;
}


HFONT FontMonospace;


static void SetEditControlFont(HWND Edit, HWND Parent)
{
	if (FontMonospace == nullptr)
	{
		HDC hdc = GetDC(Parent);
		INT dpi = GetDpi(Parent, hdc);
		INT FontSizePx = MulDiv(12 /*12pt*/, dpi, 72);

		FONT_DESC font_descs[] =
		{
			{ L"Consolas", FontSizePx },
			{ L"Courier New", FontSizePx }
		};

		FontMonospace = GetFirstMatchingFont(hdc, font_descs, sizeof(font_descs) / sizeof(font_descs[0]), nullptr);
		ReleaseDC(Parent, hdc);
		if (FontMonospace == NULL)
		{
			// Did not find an appropriate font.
			FontMonospace = (HFONT)GetStockObject(ANSI_FIXED_FONT);
		}
	}

	SendMessageW(Edit, WM_SETFONT, (WPARAM)(HFONT)FontMonospace, 0);
}


static HWND CreateEditControl(HWND Parent)
{
	SIZE ClientSize = GetClientSize(Parent);
	HWND Edit = CreateWindowExW(0, L"EDIT", nullptr, WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_WANTRETURN, 0, 0, ClientSize.cx, ClientSize.cy, Parent, nullptr, hInst, nullptr);
	assert(Edit != nullptr);

	SetEditControlFont(Edit, Parent);
	
	return Edit;
}


static BOOL OpenClipboard_ButTryABitHarder(HWND hWnd)
{
	for (int i = 0; i < 20; ++i)
	{
		// This can fail if the clipboard is currently being accessed by another application.
		if (OpenClipboard(hWnd)) return true;
		Sleep(10);
	}
	return false;
}


static void UpdateClipboard(HWND hWnd)
{
	DeleteObject(CurrentImage);
	CurrentImage = nullptr;
	free(CurrentText);
	CurrentText = nullptr;

	if (OpenClipboard_ButTryABitHarder(hWnd))
	{
		UINT KnownFormats[] =
		{
			CF_DIB,
			CF_TEXT
		};
		int ClipboardFormat = GetPriorityClipboardFormat(KnownFormats, _countof(KnownFormats));

		switch (ClipboardFormat)
		{
			case CF_DIB:
			{
				HGLOBAL ClipboardDataHandle = GetClipboardData(CF_DIB);
				if (ClipboardDataHandle)
				{
					BITMAPINFOHEADER *BitmapInfoHeader = (BITMAPINFOHEADER *)GlobalLock(ClipboardDataHandle);
					SIZE_T ClipboardDataSize = GlobalSize(ClipboardDataHandle);
					if (BitmapInfoHeader)
					{
						BITMAP BitmapDesc = {};
						HBITMAP hBitmap = CreateDIBFromPackedDIB(BitmapInfoHeader, ClipboardDataSize, &BitmapDesc);
						if (hBitmap)
						{
							CurrentImage = hBitmap;
							CurrentImageWidth = BitmapDesc.bmWidth;
							CurrentImageHeight = BitmapDesc.bmHeight;
						}
					}
				}
				break;
			}

			case CF_TEXT:
			{
				HANDLE ClipboardTextHandle = GetClipboardData(CF_UNICODETEXT);;
				if (ClipboardTextHandle != nullptr)
				{
					LPCWSTR ClipboardText = (LPCWSTR)GlobalLock(ClipboardTextHandle);
					SIZE_T ClipboardTextLength = GlobalSize(ClipboardTextHandle);
					if (ClipboardText != nullptr)
					{
						// Add a terminating 0 just in case the original data doesn't have one.
						CurrentText = (LPWSTR)malloc(sizeof(WCHAR) * (ClipboardTextLength + 1));
						if (CurrentText != nullptr)
						{
							CurrentText[ClipboardTextLength] = 0;
							memcpy(CurrentText, ClipboardText, sizeof(WCHAR) * ClipboardTextLength);
						}
					}
					GlobalUnlock(ClipboardTextHandle);
				}
				break;
			}
		}

		CloseClipboard();
	}

	UpdateCapturedContent(hWnd);
}


// Called initially, and inside UpdateClipboard after the clipboard contents have been captured.
static void UpdateCapturedContent(HWND hWnd)
{
	// Update scroll bars
	SCROLLINFO ScrollInfo = {};
	ScrollInfo.cbSize = sizeof(ScrollInfo);
	if (CurrentImage != nullptr)
	{
		// This is bugged.
		// Even if you toggle scroll bar visibility by turning the window style on and off it doesn't fix it.
		// Forcing the scroll bars to be visible *before* changing the scroll bar parameters is required to fix this.
		// If the scroll bar is not already visible at the time a SIF_DISABLENOSCROLL thing is sent, it won't show.
		ShowScrollBar(hWnd, SB_VERT, true);
		ShowScrollBar(hWnd, SB_HORZ, true);
		SIZE ClientSize = GetClientSize(hWnd);
		ScrollInfo.fMask = SIF_DISABLENOSCROLL | SIF_PAGE | SIF_RANGE;
		ScrollInfo.nPage = ClientSize.cy;
		ScrollInfo.nMax = CurrentImageHeight - 1;
		SetScrollInfo(hWnd, SB_VERT, &ScrollInfo, true);
		ScrollInfo.nPage = ClientSize.cx;
		ScrollInfo.nMax = CurrentImageWidth - 1;
		SetScrollInfo(hWnd, SB_HORZ, &ScrollInfo, true);
	}
	else
	{
		ScrollInfo.fMask = SIF_PAGE | SIF_RANGE;
		// This makes the scroll bars effectively invisible.
		SetScrollInfo(hWnd, SB_VERT, &ScrollInfo, true);
		SetScrollInfo(hWnd, SB_HORZ, &ScrollInfo, true);
	}

	// Update edit control
	if (CurrentText != nullptr)
	{
		if (CurrentEditControl == nullptr)
		{
			CurrentEditControl = CreateEditControl(hWnd);
		}
		SetWindowTextW(CurrentEditControl, CurrentText);
	}
	else
	{
		// Text is null; remove EDIT control
		if (CurrentEditControl != nullptr)
		{
			DestroyWindow(CurrentEditControl);
			CurrentEditControl = nullptr;
		}
	}

	InvalidateRect(hWnd, nullptr, true);
}


static int ScrollAmountPerLine = 10;

static BOOL Panning;
static int PanningX;
static int PanningY;

enum MONITORING_MODE
{
	MONITORING_AUTO,
	MONITORING_OFF,
	MONITORING_ONESHOT,
	MONITORING_MODE_COUNT
};

static MONITORING_MODE MonitoringMode = MONITORING_AUTO;


// Updates the menu according to TurnedOn.
static void UpdateMenuState(HWND hWnd, HMENU hMenu)
{
	if (hMenu == nullptr)
	{
		hMenu = GetMenu(hWnd);
	}

	MENUITEMINFOW MenuItemInfo = {};
	MenuItemInfo.cbSize = sizeof(MenuItemInfo);
	MenuItemInfo.fMask = MIIM_STRING;
	LPCWSTR Text;
	switch (MonitoringMode)
	{
		case MONITORING_AUTO:
			Text = L"Toggle Mode (currently AUTO)";
			break;
		case MONITORING_OFF:
		default:
			Text = L"Toggle Mode (currently OFF)";
			break;
		case MONITORING_ONESHOT:
			Text = L"Toggle Mode (currently ONE SHOT)";
			break;
	}
	MenuItemInfo.dwTypeData = (LPWSTR)Text;
	BOOL b = SetMenuItemInfoW(hMenu, IDM_TOGGLE_AUTO, false, &MenuItemInfo); assert(b);

	b = DrawMenuBar(hWnd); assert(tmp);
}


LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		case WM_CREATE:
		{
			HRESULT hr = BufferedPaintInit(); assert(SUCCEEDED(hr));

			BOOL b = AddClipboardFormatListener(hWnd); assert(b);

			HMENU Menu = CreateMenu();
			assert(Menu != nullptr);
			MENUITEMINFOW MenuItemInfo = {};
			MenuItemInfo.cbSize = sizeof(MenuItemInfo);
			MenuItemInfo.fMask = MIIM_FTYPE | MIIM_ID | MIIM_STRING;
			MenuItemInfo.fType = MFT_STRING;
			MenuItemInfo.wID = IDM_CLEAR_CLIPBOARD;
			MenuItemInfo.dwTypeData = (LPWSTR)L"Clear Clipboard Data";
			b = InsertMenuItemW(Menu, 0, false, &MenuItemInfo); assert(b);
			MenuItemInfo.wID = IDM_REFRESH;
			MenuItemInfo.dwTypeData = (LPWSTR)L"Refresh (F5)";
			b = InsertMenuItemW(Menu, 0, false, &MenuItemInfo); assert(b);
			MenuItemInfo.wID = IDM_TOGGLE_AUTO;
			MenuItemInfo.dwTypeData = (LPWSTR)L"";
			b = InsertMenuItemW(Menu, 0, false, &MenuItemInfo); assert(b);

			b = SetMenu(hWnd, Menu); assert(b);

			UpdateMenuState(hWnd, Menu);
			UpdateCapturedContent(hWnd);

			return 0;
		}

		case WM_DPICHANGED:
		{
			if (CurrentEditControl != nullptr)
			{
				HFONT OldFont = FontMonospace;
				FontMonospace = nullptr;
				SetEditControlFont(CurrentEditControl, hWnd);
				if (OldFont)
				{
					DeleteObject(OldFont);
				}
			}
			break;
		}

		case WM_KEYDOWN:
		{
			switch (wParam)
			{
				case VK_F5:
				{
					UpdateClipboard(hWnd);
					break;
				}
			}
			return 0;
		}

		case WM_COMMAND:
		{
			USHORT CommandID = LOWORD(wParam);
			switch (CommandID)
			{
				case IDM_CLEAR_CLIPBOARD:
				{
					if (OpenClipboard_ButTryABitHarder(hWnd))
					{
						EmptyClipboard();
						CloseClipboard();
						InvalidateRect(hWnd, nullptr, false);
					}
					break;
				}
				case IDM_REFRESH:
				{
					UpdateClipboard(hWnd);
					break;
				}
				case IDM_TOGGLE_AUTO:
				{
					MonitoringMode = (MONITORING_MODE)((MonitoringMode + 1) % MONITORING_MODE_COUNT);
					UpdateMenuState(hWnd, nullptr);
					break;
				}
			}
			return 0;
		}

		case WM_CLIPBOARDUPDATE:
		{
			switch (MonitoringMode)
			{
				case MONITORING_AUTO:
				case MONITORING_ONESHOT:
				{
					UpdateClipboard(hWnd);
					if (MonitoringMode == MONITORING_ONESHOT)
					{
						MonitoringMode = MONITORING_OFF;
						UpdateMenuState(hWnd, nullptr);
					}
					break;
				}
			}
			return 0;
		}

		case WM_VSCROLL:
		{
			HandleWindowMessage_Scroll(hWnd, wParam, SB_VERT, ScrollAmountPerLine, nullptr);
			return 0;
		}

		case WM_HSCROLL:
		{
			HandleWindowMessage_Scroll(hWnd, wParam, SB_HORZ, ScrollAmountPerLine, nullptr);
			return 0;
		}

		case WM_MOUSEWHEEL:
		{
			int nBar = GetKeyState(VK_SHIFT) < 0 ? SB_HORZ : SB_VERT;
			HandleWindowMessage_MouseWheel(hWnd, wParam, nBar, ScrollAmountPerLine, nullptr);
			return 0;
		}

		case WM_LBUTTONDOWN:
		{
			PanningX = GET_X_LPARAM(lParam);
			PanningY = GET_Y_LPARAM(lParam);
			Panning = true;
			return 0;
		}

		case WM_MOUSEMOVE:
		{
			if (Panning && (wParam & MK_LBUTTON))
			{
				int x = GET_X_LPARAM(lParam);
				int y = GET_Y_LPARAM(lParam);
				int dx = x - PanningX;
				int dy = y - PanningY;
				PanningX = x;
				PanningY = y;
				ScrollTo(hWnd, SB_VERT, SCROLLTO_RELATIVE, -dy, nullptr);
				ScrollTo(hWnd, SB_HORZ, SCROLLTO_RELATIVE, -dx, nullptr);
			}
			else
			{
				Panning = false;
			}
			return 0;
		}

		case WM_SIZE:
		{
			SIZE ClientSize = GetClientSize(hWnd);

			if (CurrentEditControl != nullptr)
			{
				SetWindowPos(CurrentEditControl, nullptr, 0, 0, ClientSize.cx, ClientSize.cy, 0);
			}

			if (CurrentImage != nullptr)
			{
				SCROLLINFO ScrollInfo = {};
				ScrollInfo.cbSize = sizeof(ScrollInfo);
				ScrollInfo.fMask = SIF_DISABLENOSCROLL | SIF_PAGE;
				ScrollInfo.nPage = ClientSize.cy;
				SetScrollInfo(hWnd, SB_VERT, &ScrollInfo, true);
				ScrollInfo.nPage = ClientSize.cx;
				SetScrollInfo(hWnd, SB_HORZ, &ScrollInfo, true);
			}

			return 0;
		}

		case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hdc0 = BeginPaint(hWnd, &ps);

			if (CurrentEditControl == nullptr)
			{
				HDC hdc = nullptr;
				HPAINTBUFFER PaintBuffer = BeginBufferedPaint(hdc0, &ps.rcPaint, BPBF_DIB, nullptr, &hdc);
				assert(PaintBuffer != nullptr);

				FillRect(hdc, &ps.rcPaint, (HBRUSH)GetStockObject(BLACK_BRUSH));

				if (CurrentImage != nullptr)
				{
					// Adjust for scrolling
					SCROLLINFO ScrollInfo = {};
					ScrollInfo.cbSize = sizeof(ScrollInfo);
					ScrollInfo.fMask = SIF_POS;
					GetScrollInfo(hWnd, SB_VERT, &ScrollInfo);
					int ScrollV = ScrollInfo.nPos;
					GetScrollInfo(hWnd, SB_HORZ, &ScrollInfo);
					int ScrollH = ScrollInfo.nPos;
					SetViewportOrgEx(hdc, -ScrollH, -ScrollV, nullptr);
					ps.rcPaint.top += ScrollV;
					ps.rcPaint.bottom += ScrollV;
					ps.rcPaint.left += ScrollH;
					ps.rcPaint.right += ScrollH;
			
					HDC src = CreateCompatibleDC(hdc);
					SelectObject(src, CurrentImage);
					BitBlt(hdc, 0, 0, CurrentImageWidth, CurrentImageHeight, src, 0, 0, SRCCOPY);
					DeleteDC(src);
				}

				SetViewportOrgEx(hdc, 0, 0, nullptr);
				EndBufferedPaint(PaintBuffer, true);
			}

			EndPaint(hWnd, &ps);
			return 0;
		}

		case WM_DESTROY:
		{
			BufferedPaintUnInit();
			RemoveClipboardFormatListener(hWnd);
			PostQuitMessage(0);
			return 0;
		}
	}

	return DefWindowProcW(hWnd, message, wParam, lParam);
}
