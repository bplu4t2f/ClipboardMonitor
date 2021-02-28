#include "Win32Toolbox.h"
#include <assert.h>
#include <strsafe.h>
#include <limits.h> // Required by WHEEL_PAGESCROLL -- I think that's a "bug" in the windows headers.

static BOOL GetDpi_Initialized;
static HRESULT (WINAPI *pfn_GetDpiForMonitor)(HMONITOR, DWORD, UINT*, UINT*);

// Never returns 0. In case the DPI cannot be determined for whatever Windowsy reason, returns the default 96.
INT GetDpi(HWND hWnd, HDC hdc)
{
	if (!GetDpi_Initialized)
	{
		HMODULE shcore = LoadLibraryW(L"shcore");
		// This library will never be unloaded (if we did unload it, the function pointer would become invalid anyway).
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
		if (SUCCEEDED(pfn_GetDpiForMonitor(monitor, 0, &dpi_x, &dpi_y)) && dpi_y != 0)
		{
			return dpi_y;
		}
	}

	// Need to use GetDeviceCaps.
	BOOL NeedReleaseDC = false;
	if (hdc == nullptr)
	{
		hdc = GetDC(hWnd);
		NeedReleaseDC = true;
	}
	INT LogPixelsY = GetDeviceCaps(hdc, LOGPIXELSY);
	if (LogPixelsY == 0)
	{
		// Failsafe default value. Should never happen.
		LogPixelsY = 96;
	}
	if (NeedReleaseDC)
	{
		ReleaseDC(hWnd, hdc);
		hdc = nullptr;
	}
	return LogPixelsY;
}


HFONT GetDefaultGuiFontInternal(INT dpi, BOOL *NeedDelete)
{
	HFONT StockFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
	assert(StockFont != nullptr);
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
		int tmp = GetObjectW(StockFont, sizeof(LogFont), &LogFont);
		assert(tmp != 0);
		LogFont.lfHeight = MulDiv(LogFont.lfHeight, dpi, 96);

		HFONT ScaledFont = CreateFontIndirectW(&LogFont);
		assert(ScaledFont != nullptr);
		*NeedDelete = true;
		return ScaledFont;
	}
}

HFONT GetDefaultGuiFont(DEFAULT_GUI_FONT_CACHE *Cache, INT dpi, HWND hWnd, HDC hdc)
{
	if (dpi == 0)
	{
		dpi = GetDpi(hWnd, hdc);
	}

	if (dpi == Cache->Dpi)
	{
		assert(Cache->hFont != nullptr);
		return Cache->hFont;
	}

	// DPI has changed; need to get new default font.
	if (Cache->NeedDelete)
	{
		// Clean up old font.
		DeleteObject(Cache->hFont);
		Cache->hFont = nullptr;
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
	LOGFONTW Logfont;
	BOOL Found;
};

static INT CALLBACK GetFirstMatchingFont_EnumProc(const LOGFONTW *logfont, const TEXTMETRICW *, DWORD, LPARAM lparam)
{
	GetFirstMatchingFont_LPARAM *data = (GetFirstMatchingFont_LPARAM*)lparam;
	if (!data->Found) // Actually should always be false because we stop iteration when we find anything.
	{
		data->Found = true;
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
	return nullptr;
}

INT GetTextLineHeight(TEXTMETRICW *TextMetric, BOOL WithExternalLeading)
{
	return TextMetric->tmHeight + (WithExternalLeading ? TextMetric->tmExternalLeading : 0);
}


INT GetClientWidth(HWND hWnd)
{
	RECT rc = {};
	BOOL tmp = GetClientRect(hWnd, &rc);
	assert(tmp);
	INT ClientWidth = rc.right - rc.left;
	return ClientWidth;
}

INT GetClientHeight(HWND hWnd)
{
	RECT rc = {};
	BOOL tmp = GetClientRect(hWnd, &rc);
	assert(tmp);
	INT ClientHeight = rc.bottom - rc.top;
	return ClientHeight;
}

SIZE GetClientSize(HWND hWnd)
{
	RECT rc = {};
	BOOL tmp = GetClientRect(hWnd, &rc);
	assert(tmp);
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
	INT LastScrollPos2 = ScrollInfo.nPos;

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
	SetScrollInfo(hWnd, nBar, &ScrollInfo, true);

	// Read the scroll info back because Windows sanitizes these values INTernally.
	if (!GetScrollInfo(hWnd, nBar, &ScrollInfo))
	{
		// Since GetScrollInfo doesn't reliably fail when there are no scroll bars, we must assume that
		// a second call can fail even though the first call succeeded.
		// We must invalidate the window here since we're not sure what our SetScrollInfo actually did (or if it did anything at all).
		// This happens the first time you scroll in positive direction. After that, the first GetScrollInfo will fail.
		InvalidateRect(hWnd, nullptr, true);
		return;
	}

	// NOTE: nPos will be in the range nMin .. (nMax - nPage + 1)
	//       This "+1" means that the application must most likely subtract 1 from nMax when setting the range.
	assert(ScrollInfo.nPos >= ScrollInfo.nMin);
	assert(ScrollInfo.nPos <= ScrollInfo.nMax - ScrollInfo.nPage + 1);

	int dx = 0;
	int dy = 0;
	switch (nBar)
	{
		case SB_VERT:
			dx = 0;
			dy = LastScrollPos2 - ScrollInfo.nPos;
			break;
		case SB_HORZ:
			dx = LastScrollPos2 - ScrollInfo.nPos;
			dy = 0;
			break;
	}

	if (ERROR == ScrollWindowEx(hWnd, dx, dy, ScrollRect, ScrollRect, nullptr, nullptr, SW_INVALIDATE | SW_ERASE))
	{
		// Should not be needed.
		InvalidateRect(hWnd, nullptr, true);
	}
}

void HandleWindowMessage_Scroll(HWND hWnd, WPARAM wParam, INT nBar, INT ScrollAmountPerLine, const RECT *ScrollRect)
{
	INT Request = LOWORD(wParam);
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
			if (GetScrollInfo(hWnd, nBar, &ScrollInfo))
			{
				ScrollTo(hWnd, nBar, SCROLLTO_ABSOLUTE, ScrollInfo.nTrackPos, ScrollRect);
			}
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
		INT Pages = 0;
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
		INT PixelsToScroll = MulDiv(zDelta, uScroll * ScrollAmountPerLine, WHEEL_DELTA);
		ScrollTo(hWnd, nBar, SCROLLTO_RELATIVE, -PixelsToScroll, ScrollRect);
	}
}

BOOL HandleWindowMessage_KeyDown_ForVScroll(HWND hWnd, WPARAM wParam, LPARAM lParam, INT ScrollAmountPerLine, const RECT *ScrollRect)
{
	INT vk = (INT)wParam;
	INT cRepeat = LOWORD(lParam);
	switch (vk)
	{
		case VK_UP:         ScrollTo(hWnd, SB_VERT, SCROLLTO_RELATIVE, cRepeat * -ScrollAmountPerLine, ScrollRect); return true;
		case VK_DOWN:       ScrollTo(hWnd, SB_VERT, SCROLLTO_RELATIVE, cRepeat * ScrollAmountPerLine, ScrollRect); return true;
		case VK_PRIOR:      ScrollTo(hWnd, SB_VERT, SCROLLTO_PAGE, -cRepeat, ScrollRect); return true;
		case VK_NEXT:       ScrollTo(hWnd, SB_VERT, SCROLLTO_PAGE, cRepeat, ScrollRect); return true;
		case VK_HOME:       ScrollTo(hWnd, SB_VERT, SCROLLTO_ABSOLUTE, 0, ScrollRect); return true;
		case VK_END:        ScrollTo(hWnd, SB_VERT, SCROLLTO_ABSOLUTE, MAXINT, ScrollRect); return true;
	}
	return false;
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

	if (QueryCloseRequested && *QueryCloseRequested)
	{
		return;
	}

	if (Parent != nullptr)
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
		if (QueryCloseRequested && *QueryCloseRequested)
		{
			// Normal way of closing the window. The application is expected to set this to true in its WM_CLOSE.
			break;
		}
		if (GetMessageW(&Msg, nullptr, 0, 0))
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

	if (Parent != nullptr)
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
		dpi = GetDpi(TextBox, nullptr);
	}

	// This is taken from .NET source code (TextBoxBase).
	// 3px vertical space is required between the text and the border to keep the last
	// line from being clipped.
	// This 3 pixel size was added in everett and we do this to maintain compat.
	// old everett behavior was FontHeight + [SystemInformation.BorderSize.Height * 4 + 3]
	// however the [ ] was only added if borderstyle was not none.
	INT BorderHeight = GetSystemMetricsForDpi(SM_CYBORDER, dpi);

	HFONT TextBoxFont = (HFONT)SendMessageW(TextBox, WM_GETFONT, 0, 0);
	if (TextBoxFont == nullptr)
	{
		TextBoxFont = (HFONT)GetStockObject(SYSTEM_FONT);
	}
	HDC hdc = GetDC(TextBox);
	SelectObject(hdc, TextBoxFont);
	TEXTMETRICW TextMetric = {};
	GetTextMetricsW(hdc, &TextMetric);
	ReleaseDC(TextBox, hdc);

	INT TextBoxFontHeight = TextMetric.tmHeight;
	INT TextBoxHeight = TextBoxFontHeight + BorderHeight * 4 + 3;
	return TextBoxHeight;
}


// Returns the offset, in bytes, from the start of the BITMAPINFO, to the start of the pixel data array, for a packed DIB.
INT GetPixelDataOffsetForPackedDIB(const BITMAPINFOHEADER *BitmapInfoHeader)
{
	if (BitmapInfoHeader->biSize < sizeof(BITMAPINFOHEADER))
	{
		return 0; // Not a valid BITMAPINFOHEADER. This may be an ancient BITMAPCOREHEADER, which is not supported.
	}

	INT OffsetExtra = 0;

	if (BitmapInfoHeader->biSize == sizeof(BITMAPINFOHEADER) /* 40 */)
	{
		// This is the common BITMAPINFOHEADER type. In this case, there may be bit masks following the BITMAPINFOHEADER
		// and before the actual pixel bits (does not apply if bitmap has <= 8 bpp)
		if (BitmapInfoHeader->biBitCount > 8)
		{
			if (BitmapInfoHeader->biCompression == BI_BITFIELDS)
			{
				OffsetExtra += 3 * sizeof(RGBQUAD);
			}
			else if (BitmapInfoHeader->biCompression == 6 /* BI_ALPHABITFIELDS */)
			{
				// Not widely supported, but valid.
				OffsetExtra += 4 * sizeof(RGBQUAD);
			}
		}
	}

	if (BitmapInfoHeader->biClrUsed > 0)
	{
		// We have no choice but to trust this value.
		OffsetExtra += BitmapInfoHeader->biClrUsed * sizeof(RGBQUAD);
	}
	else
	{
		// In this case, the color table contains the maximum number for the current bit count (0 if > 8bpp)
		if (BitmapInfoHeader->biBitCount <= 8)
		{
			// 1bpp: 2
			// 4bpp: 16
			// 8bpp: 256
			OffsetExtra += sizeof(RGBQUAD) << BitmapInfoHeader->biBitCount;
		}
	}

	return BitmapInfoHeader->biSize + OffsetExtra;
}


HBITMAP CreateDIBFromPackedDIB(BITMAPINFOHEADER *PackedDIB, SIZE_T PackedDIBSizeCb, BITMAP *BitmapDesc)
{
	INT PixelDataOffset = GetPixelDataOffsetForPackedDIB(PackedDIB);
	if (PixelDataOffset == 0) return nullptr;

	BYTE *PixelDataFromClipboard = (BYTE *)PackedDIB + PixelDataOffset;
	void *PixelDataNew;
	HBITMAP hBitmap = CreateDIBSection(NULL, (BITMAPINFO *)PackedDIB, DIB_RGB_COLORS, &PixelDataNew, NULL, 0);
	if (hBitmap == nullptr) return nullptr; // This will only work if the DIB format is supported by GDI. Not all valid DIB formats are supported.

	// Need to copy the data from the clipboard to the new DIBSection.
	BITMAP BitmapDescLocal; // Need this in case BitmapDesc is null.
	if (BitmapDesc == nullptr) BitmapDesc = &BitmapDescLocal;
	int tmp = GetObjectW(hBitmap, sizeof(*BitmapDesc), BitmapDesc);
	assert(tmp != 0);
	SIZE_T PixelDataBytesToCopy = (SIZE_T)BitmapDesc->bmHeight * BitmapDesc->bmWidthBytes;
	SIZE_T PixelDataBytesAvailable = PackedDIBSizeCb - PixelDataOffset;
	if (PixelDataBytesAvailable < PixelDataBytesToCopy)
	{
		// Malformed data; doesn't contain enough pixels. We'll do what we can.
		PixelDataBytesToCopy = PixelDataBytesAvailable;
	}
	memcpy(PixelDataNew, PixelDataFromClipboard, PixelDataBytesToCopy);

	return hBitmap;
}


BOOL HeapPoolEnsure(HEAP_POOL *Pool, SIZE_T Size)
{
	if (Pool->Data == nullptr || Pool->Size < Size)
	{
		// Need to allocate a [bigger] piece of memory.
		HANDLE Heap = GetProcessHeap();
		HeapFree(Heap, 0, Pool->Data);
		Pool->Data = (OUTLINETEXTMETRICW *)HeapAlloc(GetProcessHeap(), 0, Size);
		Pool->Size = Size;
		if (!Pool->Data) return false;
	}
	return true;
}

void HeapPoolFree(HEAP_POOL *Pool)
{
	HeapFree(GetProcessHeap(), 0, Pool->Data);
	Pool->Data = nullptr;
	Pool->Size = 0;
}


LPCWSTR GetFullFontFaceNameFromHDC(HDC hdc, HEAP_POOL *Pool)
{
	// Use this to get the full font name (LOGFONTW only contains the first 31 characters).
	UINT OtmRequired = GetOutlineTextMetricsW(hdc, 0, nullptr);
	if (OtmRequired == 0) return nullptr;
	if (!HeapPoolEnsure(Pool, (SIZE_T)OtmRequired + 64 /* some extra to reduce number of required allocations */)) return nullptr;
	OUTLINETEXTMETRIC *Otm = (OUTLINETEXTMETRIC *)Pool->Data;
	GetOutlineTextMetricsW(hdc, OtmRequired, Otm);
	LPCWSTR FontName = (LPCWSTR)((BYTE *)Otm + (INT_PTR)Otm->otmpFaceName /* this is actually an offset, not a string */);
	return FontName;
}


// Consider using GetFullFontFaceNameFromHDC for better performance. This avoids allocating HDC.
LPCWSTR GetFullFontFaceName(HFONT hFont, HEAP_POOL *Pool)
{
	HDC hdc = CreateCompatibleDC(nullptr);
	if (hdc == nullptr) return nullptr;
	SelectObject(hdc, hFont);
	LPCWSTR Result = GetFullFontFaceNameFromHDC(hdc, Pool);
	DeleteDC(hdc);
	return Result;
}
