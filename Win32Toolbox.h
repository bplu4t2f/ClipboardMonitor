#pragma once

#include <sdkddkver.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

struct DEFAULT_GUI_FONT_CACHE;
struct FONT_DESC;
struct HEAP_POOL;

enum SCROLLTO_MODE
{
	SCROLLTO_ABSOLUTE,
	SCROLLTO_RELATIVE,
	SCROLLTO_PAGE
};

extern INT                 GetDpi(HWND hWnd, HDC hdc);
extern HFONT               GetDefaultGuiFont(DEFAULT_GUI_FONT_CACHE *Cache, INT dpi, HWND hWnd, HDC hdc);
extern HFONT               GetDefaultGuiFontInternal(INT dpi, BOOL *NeedDelete);
extern HFONT               GetFirstMatchingFont(HDC hdc, const FONT_DESC *descs, INT num_descs, LOGFONTW *info);
extern INT                 GetClientWidth(HWND hWnd);
extern INT                 GetClientHeight(HWND hWnd);
extern SIZE                GetClientSize(HWND hWnd);
extern INT                 GetTextLineHeight(TEXTMETRICW *TextMetric, BOOL WithExternalLeading);
extern void                ScrollTo(HWND hWnd, INT nBar, SCROLLTO_MODE Mode, INT TargetPosition, const RECT *ScrollRect);
extern void                HandleWindowMessage_Scroll(HWND hWnd, WPARAM wParam, INT nBar, INT ScrollAmountPerLine, const RECT *ScrollRect);
extern void                HandleWindowMessage_MouseWheel(HWND hWnd, WPARAM wParam, INT nBar, INT ScrollAmountPerLine, const RECT *ScrollRect);
extern BOOL                HandleWindowMessage_KeyDown_ForVScroll(HWND hWnd, WPARAM wParam, LPARAM lParam, INT ScrollAmountPerLine, const RECT *ScrollRect);
extern INT                 StrlenMax(LPCWSTR str, INT cchMax);
extern void                ShowWindowModal(HWND hWnd, BOOL *QueryCloseRequested);
extern INT                 GetDefaultSinglelineEditBoxHeight(HWND TextBox, INT dpi);
extern INT                 GetPixelDataOffsetForPackedDIB(const BITMAPINFOHEADER *BitmapInfoHeader);
extern HBITMAP             CreateDIBFromPackedDIB(BITMAPINFOHEADER *PackedDIB, SIZE_T PackedDIBSizeCb, BITMAP *BitmapDesc);
extern BOOL                HeapPoolEnsure(HEAP_POOL *Pool, SIZE_T Size);
extern void                HeapPoolFree(HEAP_POOL *Pool);
extern LPCWSTR             GetFullFontFaceNameFromHDC(HDC hdc, HEAP_POOL *Pool);
extern LPCWSTR             GetFullFontFaceName(HFONT hFont, HEAP_POOL *Pool);

struct DEFAULT_GUI_FONT_CACHE
{
	HFONT hFont;
	INT Dpi;
	BOOL NeedDelete;
};

struct FONT_DESC
{
	const WCHAR *Name;
	int Height;
	// Could add weight, italic, etc...
};

struct HEAP_POOL
{
	void *Data;
	size_t Size;
};

// This can be used with DialogBoxIndirectParamW, or CreateDialogIndirectParamW, to create an empty dialog (DLGTEMPLATE won't cut it).
// You'll have to initialize the contents of the dialog in your WM_INITDIALOG. That includes the caption text (use SetWindowText on the hDlg).
struct DLGTEMPLATE_EMPTY
{
	// Recommended:
	//  WS_CAPTION | WS_SYSMENU | WS_POPUP
	// Consider: 
	//  WS_THICKFRAME (makes window resizable)
	DWORD style;
	DWORD dwExtendedStyle;
	WORD Reserved_cdit_MustBe0;
	short x;
	short y;
	short cx;
	short cy;
	WCHAR Reserved_MenuName_MustBe0[1];
	WCHAR Reserved_WindowClassName_MustBe0[1];
	WCHAR Reserved_WindowTitle_MustBe0[1];
};
