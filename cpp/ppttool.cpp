#include <windows.h>
#include <vector>
#include <string>

#define COLOR_BORDER RGB(153, 153, 153)
#define COLOR_BG RGB(255, 255, 255)
#define COLOR_ACTIVE RGB(230, 230, 230)
#define COLOR_TEXT RGB(0, 0, 0)
#define COLOR_SEP RGB(153, 153, 153)
#define COLOR_MAGIC RGB(255, 0, 255)

// 虚拟键码
#define VK_PRIOR 0x21
#define VK_NEXT 0x22
#define VK_ESCAPE 0x1B
#define VK_CONTROL 0x11
#define VK_M 0x4D
#define VK_P 0x50
#define VK_E 0x45
#define VK_UP 0x26
#define VK_DOWN 0x28
#define VK_LEFT 0x25
#define VK_RIGHT 0x27

// 全局变量
HWND g_pptHwnd = NULL;
HWND g_verticalWnd = NULL;
HWND g_horizontalWnd = NULL;
bool g_pptPlaying = false;
std::vector<bool> g_buttonsPressed;
std::vector<RECT> g_buttonRects;
std::vector<std::string> g_buttonTexts;
bool g_prevPressed = false;
bool g_nextPressed = false;

// 查找PPT窗口
HWND FindPptWindow() {
	HWND result = NULL;
	EnumWindows([](HWND hwnd, LPARAM lparam) -> BOOL {
		if (IsWindowVisible(hwnd)) {
			char title[256];
			GetWindowTextA(hwnd, title, 256);
			if (strstr(title, "幻灯片放映") || strstr(title, "Slide Show")) {
				*(HWND*)lparam = hwnd;
				return FALSE;
			}
		}
		return TRUE;
	}, (LPARAM)&result);
	return result;
}

// 激活PPT窗口
bool ActivatePpt() {
	if (g_pptHwnd && IsWindow(g_pptHwnd)) {
		ShowWindow(g_pptHwnd, SW_SHOW);
		SetForegroundWindow(g_pptHwnd);
		return true;
	}
	g_pptHwnd = FindPptWindow();
	if (g_pptHwnd) {
		ShowWindow(g_pptHwnd, SW_SHOW);
		SetForegroundWindow(g_pptHwnd);
		return true;
	}
	return false;
}

// 发送按键
void SendKey(BYTE vkCode) {
	if (ActivatePpt()) {
		keybd_event(vkCode, 0, 0, 0);
		keybd_event(vkCode, 0, KEYEVENTF_KEYUP, 0);
	}
}

void SendCtrlKey(BYTE vkCode) {
	if (ActivatePpt()) {
		keybd_event(VK_CONTROL, 0, 0, 0);
		keybd_event(vkCode, 0, 0, 0);
		keybd_event(vkCode, 0, KEYEVENTF_KEYUP, 0);
		keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
	}
}

void DrawRoundedRect(HDC hdc, int x, int y, int w, int h, int r, COLORREF color) {
	HBRUSH brush = CreateSolidBrush(color);
	HPEN pen = CreatePen(PS_NULL, 0, color);
	HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, brush);
	HPEN oldPen = (HPEN)SelectObject(hdc, pen);
	RoundRect(hdc, x, y, x + w, y + h, r * 2, r * 2);
	SelectObject(hdc, oldBrush);
	SelectObject(hdc, oldPen);
	DeleteObject(brush);
	DeleteObject(pen);
}

void DrawText(HDC hdc, int x, int y, int w, int h, const char* text, int fontSize) {
	HFONT font = CreateFontA(fontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
							 DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
							 DEFAULT_QUALITY, DEFAULT_PITCH, "Microsoft YaHei");
	HFONT oldFont = (HFONT)SelectObject(hdc, font);
	SetTextColor(hdc, COLOR_TEXT);
	SetBkMode(hdc, TRANSPARENT);
	RECT rect = {x, y, x + w, y + h};
	DrawTextA(hdc, text, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
	SelectObject(hdc, oldFont);
	DeleteObject(font);
}

// 垂直窗口过程
LRESULT CALLBACK VerticalWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	static int btnHeight = 60;
	
	switch (msg) {
		case WM_CREATE: {
		SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_LAYERED);
		SetLayeredWindowAttributes(hwnd, COLOR_MAGIC, 127, LWA_COLORKEY | LWA_ALPHA);
		break;
	}
		
	case WM_ERASEBKGND:
		return 1;
		
		case WM_PAINT: {
			PAINTSTRUCT ps;
			HDC hdc = BeginPaint(hwnd, &ps);
			RECT rect;
			GetClientRect(hwnd, &rect);
			int w = rect.right;
			int h = rect.bottom;
			
			HBRUSH magicBrush = CreateSolidBrush(COLOR_MAGIC);
			FillRect(hdc, &rect, magicBrush);
			DeleteObject(magicBrush);
			
			DrawRoundedRect(hdc, 0, 0, w, h, 8, COLOR_BORDER);
			DrawRoundedRect(hdc, 1, 1, w - 2, h - 2, 7, COLOR_BG);
			
			COLORREF prevColor = g_prevPressed ? COLOR_ACTIVE : COLOR_BG;
			DrawRoundedRect(hdc, 1, 1, w - 2, btnHeight - 1, 7, prevColor);
			DrawText(hdc, 1, 1, w - 2, btnHeight - 1, "△", 40);
			
			HPEN pen = CreatePen(PS_SOLID, 1, COLOR_SEP);
			HPEN oldPen = (HPEN)SelectObject(hdc, pen);
			MoveToEx(hdc, 1, btnHeight, NULL);
			LineTo(hdc, w - 1, btnHeight);
			SelectObject(hdc, oldPen);
			DeleteObject(pen);
			
			COLORREF nextColor = g_nextPressed ? COLOR_ACTIVE : COLOR_BG;
			DrawRoundedRect(hdc, 1, btnHeight + 1, w - 2, btnHeight - 1, 7, nextColor);
			DrawText(hdc, 1, btnHeight + 1, w - 2, btnHeight - 1, "▽", 40);
			
			EndPaint(hwnd, &ps);
			break;
		}
		
		case WM_LBUTTONDOWN: {
			int y = HIWORD(lParam);
			RECT rect;
			GetClientRect(hwnd, &rect);
			int btnH = rect.bottom / 2;
			
			if (y < btnH) {
				g_prevPressed = true;
				InvalidateRect(hwnd, NULL, TRUE);
				SendKey(VK_PRIOR);
			} else {
				g_nextPressed = true;
				InvalidateRect(hwnd, NULL, TRUE);
				SendKey(VK_NEXT);
			}
			break;
		}
		
		case WM_LBUTTONUP: {
			g_prevPressed = false;
			g_nextPressed = false;
			InvalidateRect(hwnd, NULL, TRUE);
			break;
		}
		
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}

// 水平窗口过程
LRESULT CALLBACK HorizontalWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_CREATE: {
		SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_LAYERED);
		SetLayeredWindowAttributes(hwnd, COLOR_MAGIC, 127, LWA_COLORKEY | LWA_ALPHA);
		
		g_buttonsPressed.resize(5, false);
		g_buttonRects.resize(5);
		g_buttonTexts.push_back("鼠标");
		g_buttonTexts.push_back("画笔");
		g_buttonTexts.push_back("橡皮");
		g_buttonTexts.push_back("清空");
		g_buttonTexts.push_back("结束放映");
		break;
	}
		
	case WM_ERASEBKGND:
		return 1;
		
		case WM_SIZE: {
			RECT rect;
			GetClientRect(hwnd, &rect);
			int w = rect.right;
			int h = rect.bottom;
			int btnWidth = (w - 2 - 4) / 5;
			
			for (int i = 0; i < 5; i++) {
				g_buttonRects[i].left = 1 + i * (btnWidth + 1);
				g_buttonRects[i].top = 1;
				g_buttonRects[i].right = g_buttonRects[i].left + btnWidth;
				g_buttonRects[i].bottom = h - 1;
			}
			break;
		}
		
		case WM_PAINT: {
			PAINTSTRUCT ps;
			HDC hdc = BeginPaint(hwnd, &ps);
			RECT rect;
			GetClientRect(hwnd, &rect);
			int w = rect.right;
			int h = rect.bottom;
			
			HBRUSH magicBrush = CreateSolidBrush(COLOR_MAGIC);
			FillRect(hdc, &rect, magicBrush);
			DeleteObject(magicBrush);
			
			DrawRoundedRect(hdc, 0, 0, w, h, 8, COLOR_BORDER);
			DrawRoundedRect(hdc, 1, 1, w - 2, h - 2, 7, COLOR_BG);
			
			for (int i = 0; i < 5; i++) {
				COLORREF btnColor = g_buttonsPressed[i] ? COLOR_ACTIVE : COLOR_BG;
				DrawRoundedRect(hdc, g_buttonRects[i].left, g_buttonRects[i].top,
								g_buttonRects[i].right - g_buttonRects[i].left,
								g_buttonRects[i].bottom - g_buttonRects[i].top, 7, btnColor);
				DrawText(hdc, g_buttonRects[i].left, g_buttonRects[i].top,
						 g_buttonRects[i].right - g_buttonRects[i].left,
						 g_buttonRects[i].bottom - g_buttonRects[i].top, 
						 g_buttonTexts[i].c_str(), 20);
				
				if (i < 4) {
					HPEN pen = CreatePen(PS_SOLID, 1, COLOR_SEP);
					HPEN oldPen = (HPEN)SelectObject(hdc, pen);
					MoveToEx(hdc, g_buttonRects[i].right, 1, NULL);
					LineTo(hdc, g_buttonRects[i].right, h - 1);
					SelectObject(hdc, oldPen);
					DeleteObject(pen);
				}
			}
			
			EndPaint(hwnd, &ps);
			break;
		}
		
		case WM_LBUTTONDOWN: {
			int x = LOWORD(lParam);
			int y = HIWORD(lParam);
			
			for (int i = 0; i < 5; i++) {
				if (x >= g_buttonRects[i].left && x <= g_buttonRects[i].right &&
					y >= g_buttonRects[i].top && y <= g_buttonRects[i].bottom) {
					g_buttonsPressed[i] = true;
					InvalidateRect(hwnd, NULL, TRUE);
					
					// 连续快捷键映射（鼠标点击）
					switch(i) {
						case 0: // 鼠标
						SendCtrlKey(VK_M);
						SendCtrlKey(VK_M);
						break;
						case 1: // 画笔
						SendCtrlKey(VK_M);
						SendCtrlKey(VK_M);
						SendCtrlKey(VK_P);
						break;
						case 2: // 橡皮
						SendCtrlKey(VK_E);
						break;
						case 3: // 清空
						SendKey(VK_E);
						break;
						case 4: // 结束放映
						SendKey(VK_ESCAPE);
						SendKey(VK_ESCAPE);
						break;
					}
					break;
				}
			}
			break;
		}
		
		case WM_LBUTTONUP: {
			for (int i = 0; i < 5; i++) {
				g_buttonsPressed[i] = false;
			}
			InvalidateRect(hwnd, NULL, TRUE);
			break;
		}
		
		case WM_KEYDOWN: {
			// 键盘快捷键映射（连续快捷键）
			switch(wParam) {
			case VK_UP:
			case VK_LEFT:
				SendKey(VK_PRIOR);
				return 0;
			case VK_DOWN:
			case VK_RIGHT:
				SendKey(VK_NEXT);
				return 0;
			case 'M':
			case 'm':
				SendCtrlKey(VK_M);
				SendCtrlKey(VK_M);
				return 0;
			case 'P':
			case 'p':
				SendCtrlKey(VK_M);
				SendCtrlKey(VK_M);
				SendCtrlKey(VK_P);
				return 0;
			case 'E':
			case 'e':
				SendCtrlKey(VK_E);
				return 0;
			case 'C':
			case 'c':
				SendKey(VK_E);
				return 0;
			case VK_ESCAPE:
				SendKey(VK_ESCAPE);
				SendKey(VK_ESCAPE);
				return 0;
			}
			break;
		}
		
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}

// 创建窗口
HWND CreateToolWindow(HINSTANCE hInst, const char* className, WNDPROC wndProc, int width, int height, int x, int y) {
	WNDCLASSEXA wc = {0};
	wc.cbSize = sizeof(WNDCLASSEXA);
	wc.lpfnWndProc = wndProc;
	wc.hInstance = hInst;
	wc.hCursor = LoadCursor(NULL, IDC_HAND);
	wc.hbrBackground = (HBRUSH)GetStockObject(NULL_BRUSH);
	wc.lpszClassName = className;
	RegisterClassExA(&wc);
	
	HWND hwnd = CreateWindowExA(
								WS_EX_NOACTIVATE | WS_EX_TOPMOST,
								className, "", WS_POPUP,
								x, y, width + 2, height,
								NULL, NULL, hInst, NULL
								);
	
	HMENU menu = GetSystemMenu(hwnd, FALSE);
	if (menu) RemoveMenu(menu, SC_CLOSE, MF_BYCOMMAND);
	
	ShowWindow(hwnd, SW_HIDE);
	return hwnd;
}

// 监控PPT状态线程
DWORD WINAPI MonitorThread(LPVOID lpParam) {
	bool lastState = false;
	
	while (true) {
		HWND pptHwnd = FindPptWindow();
		bool state = (pptHwnd != NULL);
		
		if (state != lastState) {
			lastState = state;
			if (state) {
				g_pptHwnd = pptHwnd;
				ShowWindow(g_verticalWnd, SW_SHOW);
				ShowWindow(g_horizontalWnd, SW_SHOW);
			} else {
				ShowWindow(g_verticalWnd, SW_HIDE);
				ShowWindow(g_horizontalWnd, SW_HIDE);
			}
		}
		Sleep(1000);
	}
	return 0;
}

// 主函数
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpCmdLine, int nCmdShow) {
	int screenWidth = GetSystemMetrics(SM_CXSCREEN);
	int screenHeight = GetSystemMetrics(SM_CYSCREEN);
	
	g_verticalWnd = CreateToolWindow(hInst, "VerticalWindow", VerticalWndProc, 60, 121, 
									 screenWidth - 62, (screenHeight - 121) / 2);
	
	g_horizontalWnd = CreateToolWindow(hInst, "HorizontalWindow", HorizontalWndProc, 504, 45,
									   (screenWidth - 506) / 2, screenHeight - 45);
	
	CreateThread(NULL, 0, MonitorThread, NULL, 0, NULL);
	
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	
	return 0;
}
