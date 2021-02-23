#include "Win32Toolbox.h"
#include <assert.h>
#include <strsafe.h>
#include <limits.h> // Required by WHEEL_PAGESCROLL -- I think that's a "bug" in the windows headers.

static bool GetDpi_Initialized;
static HRESULT (WINAPI *pfn_GetDpiForMonitor)(HMONITOR, DWORD, UINT*, UINT*);

INT GetDpi(HWND hWnd, HDC hdc)
{
	if (!GetDpi_Initialized)
	{
		HMODULE shcore = LoadLibraryW(L"shcore");
		if (shcore)
		{
			pfn_GetDpiForMonitor = (HRESULT (WINAPI *)(HMONITOR, DWORD, UINT*, UINT*))GetProcAddress(shcore, "GetDpiForMonitor");
		}
		GetDpi_Initialized = true;
	}

	if (pfn_GetDpiForMonitor != nullptr)
	{
		HMONITOR monitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTOPRIMARY);
		UINT dpi_x = 0;
		UINT dpi_y = 0;
		if (SUCCEEDED(pfn_GetDpiForMonitor(monitor, 0, &dpi_x, &dpi_y)))
		{
			return dpi_x;
		}
	}

	// Need to use GetDeviceCaps.
	bool NeedReleaseDC = false;
	if (hdc == NULL)
	{
		hdc = GetDC(hWnd);
		NeedReleaseDC = true;
	}
	int LogPixelsX = GetDeviceCaps(hdc, LOGPIXELSX);
	if (NeedReleaseDC)
	{
		ReleaseDC(hWnd, hdc);
		hdc = NULL;
	}
	return LogPixelsX;
}


HFONT GetDefaultGuiFontInternal(INT dpi, BOOL *NeedDelete)
{
	HFONT StockFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
	assert(StockFont != NULL);
	if (dpi == 96)
	{
		// No adjustment needed
		*NeedDelete = false;
		return StockFont;
	}
	else
	{
		// DPI scaling needed.
		LOGFONTW LogFont = {};
		assert(NULL != GetObjectW(StockFont, sizeof(LogFont), &LogFont));
		LogFont.lfHeight = MulDiv(LogFont.lfHeight, dpi, 96);

		HFONT ScaledFont = CreateFontIndirectW(&LogFont);
		assert(ScaledFont != NULL);
		*NeedDelete = true;
		return ScaledFont;
	}
}

HFONT GetDefaultGuiFont(DEFAULT_GUI_FONT_CACHE *Cache, INT dpi, HWND hWnd, HDC hdc)
{
	if (dpi == 0)
	{
		dpi = GetDpi(hWnd, hdc);
		assert(dpi != 0);
	}

	if (dpi == Cache->Dpi)
	{
		assert(Cache->hFont != NULL);
		return Cache->hFont;
	}

	// DPI has changed; need to get new default font.
	if (Cache->NeedDelete)
	{
		// Clean up old font.
		DeleteObject(Cache->hFont);
		Cache->hFont = NULL;
		Cache->Dpi = 0;
		Cache->NeedDelete = false;
	}

	BOOL NeedDelete;
	Cache->hFont = GetDefaultGuiFontInternal(dpi, &NeedDelete);
	Cache->NeedDelete = NeedDelete;
	Cache->Dpi = dpi;

	return Cache->hFont;
}

struct GetFirstMatchingFont_LPARAM
{
	BOOL Found;
	LOGFONTW Logfont;
};

static int CALLBACK GetFirstMatchingFont_EnumProc(const LOGFONTW *logfont, const TEXTMETRICW *, DWORD, LPARAM lparam)
{
	GetFirstMatchingFont_LPARAM *data = (GetFirstMatchingFont_LPARAM*)lparam;
	if (!data->Found) // Actually should always be FALSE because we stop iteration when we find anything.
	{
		data->Found = TRUE;
		data->Logfont = *logfont;
	}
	return 0; // Return 0 means "break". We don't want to continue once the first font was found.
}

HFONT GetFirstMatchingFont(HDC hdc, const FONT_DESC *descs, INT num_descs, LOGFONTW *info)
{
	// Iterate through fonts using EnumFontFamiliesW, until we find the first one that matches a name in descs.
	for (INT i = 0; i < num_descs; ++i)
	{
		auto it = descs + i;
		GetFirstMatchingFont_LPARAM lparam = {};
		LOGFONTW probe = {};
		probe.lfCharSet = DEFAULT_CHARSET;
		StringCbCopyW(probe.lfFaceName, sizeof(probe.lfFaceName), it->Name);
		EnumFontFamiliesExW(hdc, &probe, GetFirstMatchingFont_EnumProc, (LPARAM)&lparam, 0);
		if (lparam.Found)
		{
			// We have found at least one matching font, and the LOGFONTW is in lparam.Logfont. Use that to create a HFONT.
			lparam.Logfont.lfHeight = it->Height;
			// TODO maybe we want to do that DPI awareness thing somehow...
			lparam.Logfont.lfWidth = 0;
			HFONT hfont = CreateFontIndirectW(&lparam.Logfont);
			if (info != nullptr)
			{
				*info = lparam.Logfont;
			}
			return hfont;
		}
	}
	// No matching font found.
	return NULL;
}

INT GetTextLineHeight(TEXTMETRICW *TextMetric, BOOL WithExternalLeading)
{
	return TextMetric->tmHeight + (WithExternalLeading ? TextMetric->tmExternalLeading : 0);
}


INT GetClientWidth(HWND hWnd)
{
	RECT rc = {};
	assert(GetClientRect(hWnd, &rc));
	int ClientWidth = rc.right - rc.left;
	return ClientWidth;
}

INT GetClientHeight(HWND hWnd)
{
	RECT rc = {};
	assert(GetClientRect(hWnd, &rc));
	int ClientHeight = rc.bottom - rc.top;
	return ClientHeight;
}

SIZE GetClientSize(HWND hWnd)
{
	RECT rc = {};
	assert(GetClientRect(hWnd, &rc));
	SIZE size = {};
	size.cx = rc.right - rc.left;
	size.cy = rc.bottom - rc.top;
	return size;
}

void ScrollTo(HWND hWnd, INT nBar, SCROLLTO_MODE Mode, INT TargetPosition, const RECT *ScrollRect)
{
	SCROLLINFO ScrollInfo = {};
	ScrollInfo.cbSize = sizeof(ScrollInfo);
	ScrollInfo.fMask = SIF_POS | SIF_PAGE;
	// Need to get old scroll position first, so that we know the amount by which the window needs to scroll.
	if (!GetScrollInfo(hWnd, nBar, &ScrollInfo))
	{
		// This *sometimes* happens if there are no scroll bars to begin with. In that case, we can't scroll.
		// NOTE: It never happens if you scroll in negative direction, but always if you scroll in positive direction.
		return;
	}
	int LastScrollPos2 = ScrollInfo.nPos;

	switch (Mode)
	{
		case SCROLLTO_ABSOLUTE:
			ScrollInfo.nPos = TargetPosition;
			break;
		case SCROLLTO_RELATIVE:
			ScrollInfo.nPos += TargetPosition;
			break;
		case SCROLLTO_PAGE:
			// Here, TargetPosition means number of pages.
			ScrollInfo.nPos += ScrollInfo.nPage * TargetPosition;
			break;
	}

	ScrollInfo.fMask = SIF_POS;
	SetScrollInfo(hWnd, nBar, &ScrollInfo, TRUE);

	// Read the scroll info back because Windows sanitizes these values internally.
	if (!GetScrollInfo(hWnd, nBar, &ScrollInfo))
	{
		// Since GetScrollInfo doesn't reliably fail when there are no scroll bars, we must assume that
		// a second call can fail even though the first call succeeded.
		// We must invalidate the window here since we're not sure what our SetScrollInfo actually did (or if it did anything at all).
		// This happens the first time you scroll in positive direction. After that, the first GetScrollInfo will fail.
		InvalidateRect(hWnd, NULL, TRUE);
		return;
	}

	switch (nBar)
	{
		case SB_VERT:
			assert(ERROR != ScrollWindowEx(hWnd, 0, LastScrollPos2 - ScrollInfo.nPos, ScrollRect, ScrollRect, NULL, nullptr, SW_INVALIDATE | SW_ERASE));
			break;
		case SB_HORZ:
			assert(ERROR != ScrollWindowEx(hWnd, LastScrollPos2 - ScrollInfo.nPos, 0, ScrollRect, ScrollRect, NULL, nullptr, SW_INVALIDATE | SW_ERASE));
			break;
	}
	
	// Should not be needed; only here for certain debugging scenarios.
	//InvalidateRect(hWnd, NULL, TRUE);
}

void HandleWindowMessage_Scroll(HWND hWnd, WPARAM wParam, INT nBar, INT ScrollAmountPerLine, const RECT *ScrollRect)
{
	int Request = LOWORD(wParam);
	switch (Request)
	{
		case SB_LINEUP:        ScrollTo(hWnd, nBar, SCROLLTO_RELATIVE, -ScrollAmountPerLine, ScrollRect); break;
		case SB_LINEDOWN:      ScrollTo(hWnd, nBar, SCROLLTO_RELATIVE, ScrollAmountPerLine, ScrollRect); break;
		case SB_PAGEUP:        ScrollTo(hWnd, nBar, SCROLLTO_PAGE, -1, ScrollRect); break;
		case SB_PAGEDOWN:      ScrollTo(hWnd, nBar, SCROLLTO_PAGE, 1, ScrollRect); break;
		case SB_TOP:           ScrollTo(hWnd, nBar, SCROLLTO_ABSOLUTE, 0, ScrollRect); break;
		case SB_BOTTOM:        ScrollTo(hWnd, nBar, SCROLLTO_ABSOLUTE, MAXINT, ScrollRect); break;
		case SB_THUMBTRACK:
		{
			// The correct way to handle this is to first get the current
			// thumb position, and then scroll to that position.
			// Do not trust the value given with the message, it's only 16 bits.
			SCROLLINFO ScrollInfo = {};
			ScrollInfo.cbSize = sizeof(ScrollInfo);
			ScrollInfo.fMask = SIF_TRACKPOS;
			assert(GetScrollInfo(hWnd, nBar, &ScrollInfo));
			ScrollTo(hWnd, nBar, SCROLLTO_ABSOLUTE, ScrollInfo.nTrackPos, ScrollRect);
			break;
		}
	}
}

void HandleWindowMessage_MouseWheel(HWND hWnd, WPARAM wParam, INT nBar, INT ScrollAmountPerLine, const RECT *ScrollRect)
{
	INT16 zDelta = (INT16)HIWORD(wParam);

	UINT uScroll = 0;
	if (!SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &uScroll, 0))
	{
		uScroll = 3; // default value
	}

	if (uScroll == 0) return;

	// NOTE: A positive PixelsToScroll value means the mouse wheel was scrolled upwards, which in turn should decrease the
	//       current scroll bar position.
	if (uScroll == WHEEL_PAGESCROLL)
	{
		// Scroll a page
		int Pages = 0;
		if (zDelta > 0)
		{
			Pages = -1;
		}
		else if (zDelta < 0)
		{
			Pages = 1;
		}
		ScrollTo(hWnd, nBar, SCROLLTO_PAGE, Pages, ScrollRect);
	}
	else
	{
		int PixelsToScroll = MulDiv(zDelta, uScroll * ScrollAmountPerLine, WHEEL_DELTA);
		ScrollTo(hWnd, nBar, SCROLLTO_RELATIVE, -PixelsToScroll, ScrollRect);
	}
}

BOOL HandleWindowMessage_KeyDown_ForVScroll(HWND hWnd, WPARAM wParam, LPARAM lParam, INT ScrollAmountPerLine, const RECT *ScrollRect)
{
	int vk = (int)wParam;
	int cRepeat = LOWORD(lParam);
	switch (vk)
	{
		case VK_UP:         ScrollTo(hWnd, SB_VERT, SCROLLTO_RELATIVE, cRepeat * -ScrollAmountPerLine, ScrollRect); return TRUE;
		case VK_DOWN:       ScrollTo(hWnd, SB_VERT, SCROLLTO_RELATIVE, cRepeat * ScrollAmountPerLine, ScrollRect); return TRUE;
		case VK_PRIOR:      ScrollTo(hWnd, SB_VERT, SCROLLTO_PAGE, -cRepeat, ScrollRect); return TRUE;
		case VK_NEXT:       ScrollTo(hWnd, SB_VERT, SCROLLTO_PAGE, cRepeat, ScrollRect); return TRUE;
		case VK_HOME:       ScrollTo(hWnd, SB_VERT, SCROLLTO_ABSOLUTE, 0, ScrollRect); return TRUE;
		case VK_END:        ScrollTo(hWnd, SB_VERT, SCROLLTO_ABSOLUTE, MAXINT, ScrollRect); return TRUE;
	}
	return FALSE;
}

INT StrlenMax(LPCWSTR str, INT cchMax)
{
	size_t cchLength = 0;
	if (SUCCEEDED(StringCchLengthW(str, cchMax, &cchLength)))
	{
		return cchLength;
	}
	else
	{
		return cchMax;
	}
}


void ShowWindowModal(HWND hWnd, BOOL *QueryCloseRequested)
{
	HWND Parent = GetParent(hWnd);

	assert(QueryCloseRequested != nullptr);
	if (*QueryCloseRequested)
	{
		return;
	}

	if (Parent != NULL)
	{
		EnableWindow(Parent, false);
	}

	ShowWindow(hWnd, SW_SHOW);

	MSG Msg;
	while (true)
	{
		if (!IsWindow(hWnd))
		{
			// This prevents us from being stuck in the modal loop if the window is destroyed unexpectedly.
			break;
		}
		if (*QueryCloseRequested)
		{
			// Normal way of closing the window. The application is expected to set this to true in its WM_CLOSE.
			break;
		}
		if (GetMessageW(&Msg, NULL, 0, 0))
		{
			if (!IsDialogMessageW(hWnd, &Msg))
			{
				TranslateMessage(&Msg);
				DispatchMessageW(&Msg);
			}
		}
		else
		{
			// WM_QUIT; repost and break. This is supposed to tear down the entire application.
			PostQuitMessage(Msg.wParam);
			break;
		}
	}

	if (Parent != NULL)
	{
		EnableWindow(Parent, true);
		// This is a failsafe measure working around certain Windows "features".
		//SetActiveWindow(Parent);
	}

	DestroyWindow(hWnd);
}

INT GetDefaultSinglelineEditBoxHeight(HWND TextBox, INT dpi)
{
	if (dpi == 0)
	{
		dpi = GetDpi(TextBox, NULL);
	}

	// This is taken from .NET source code (TextBoxBase).
	// 3px vertical space is required between the text and the border to keep the last
	// line from being clipped.
	// This 3 pixel size was added in everett and we do this to maintain compat.
	// old everett behavior was FontHeight + [SystemInformation.BorderSize.Height * 4 + 3]
	// however the [ ] was only added if borderstyle was not none.
	int BorderHeight = GetSystemMetricsForDpi(SM_CYBORDER, dpi);

	HFONT TextBoxFont = (HFONT)SendMessageW(TextBox, WM_GETFONT, 0, 0);
	if (TextBoxFont == NULL)
	{
		TextBoxFont = (HFONT)GetStockObject(SYSTEM_FONT);
	}
	HDC hdc = GetDC(TextBox);
	SelectObject(hdc, TextBoxFont);
	TEXTMETRICW TextMetric = {};
	GetTextMetricsW(hdc, &TextMetric);
	ReleaseDC(TextBox, hdc);

	int TextBoxFontHeight = TextMetric.tmHeight;
	int TextBoxHeight = TextBoxFontHeight + BorderHeight * 4 + 3;
	return TextBoxHeight;
}
