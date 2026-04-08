#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <vector>
#include <string>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <cctype>
#include <cmath>

// 颜色定义
#define COLOR_BORDER RGB(153, 153, 153)
#define COLOR_BG RGB(255, 255, 255)
#define COLOR_ACTIVE RGB(230, 230, 230)
#define COLOR_TEXT RGB(0, 0, 0)
#define COLOR_SEP RGB(153, 153, 153)
#define COLOR_MAGIC RGB(255, 0, 255)
#define COLOR_SLIDER_BG RGB(224, 224, 224)      // 滑块背景灰色
#define COLOR_SLIDER_FILL RGB(0, 120, 215)      // Win10 蓝色填充
#define COLOR_SLIDER_THUMB RGB(0, 120, 215)     // 滑块按钮蓝色

// 虚拟键码
#define VK_PRIOR 0x21
#define VK_NEXT 0x22
#define VK_ESCAPE 0x1B
#define VK_CONTROL 0x11
#define VK_A 0x41
#define VK_P 0x50
#define VK_E 0x45

// 配置结构体
struct Config {
	bool enableLeftNav = true;
	bool enableRightNav = true;
	bool enableToolbar = true;
	bool enableLeftBottomNav = true;
	bool enableRightBottomNav = true;
};

Config g_config;
HWND g_pptHwnd = NULL;
HWND g_leftWnd = NULL;
HWND g_rightWnd = NULL;
HWND g_leftBottomWnd = NULL;
HWND g_rightBottomWnd = NULL;
HWND g_horizontalWnd = NULL;
std::vector<bool> g_buttonsPressed;
std::vector<RECT> g_buttonRects;
std::vector<std::string> g_buttonTexts;
bool g_leftPrevPressed = false;
bool g_leftNextPressed = false;
bool g_rightPrevPressed = false;
bool g_rightNextPressed = false;
bool g_leftBottomPrevPressed = false;
bool g_leftBottomNextPressed = false;
bool g_rightBottomPrevPressed = false;
bool g_rightBottomNextPressed = false;

// 音量窗口句柄和相关变量
HWND g_volumeWnd = NULL;
int g_currentVolume = 50;
bool g_dragging = false;
int g_sliderWidth = 140;
int g_trackHeight = 8;      // 轨道高度
// [修改点 3] 滑块按钮宽度调整为原来的一半 (16 -> 8)
int g_thumbWidth = 8;       
int g_thumbHeight = 20;     // 按钮高度
int g_sliderX = 10;
int g_sliderY = 0;          // 动态计算
int g_thumbX = 0;

// =========================== 音量控制 API ===========================
IAudioEndpointVolume* g_pVolume = NULL;

bool InitVolumeControl() {
	HRESULT hr;
	IMMDeviceEnumerator* pEnumerator = NULL;
	IMMDevice* pDevice = NULL;
	
	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	if (FAILED(hr)) return false;
	
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
						  CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
						  (void**)&pEnumerator);
	if (FAILED(hr)) return false;
	
	hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
	if (FAILED(hr)) {
		pEnumerator->Release();
		return false;
	}
	
	hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL,
						   NULL, (void**)&g_pVolume);
	pDevice->Release();
	pEnumerator->Release();
	
	return SUCCEEDED(hr);
}

void UninitVolumeControl() {
	if (g_pVolume) {
		g_pVolume->Release();
		g_pVolume = NULL;
	}
	CoUninitialize();
}

int GetSystemVolume() {
	if (!g_pVolume) return 50;
	float level;
	g_pVolume->GetMasterVolumeLevelScalar(&level);
	return (int)(level * 100 + 0.5);
}

void SetSystemVolume(int percent) {
	if (!g_pVolume) return;
	float level = std::max(0, std::min(100, percent)) / 100.0f;
	g_pVolume->SetMasterVolumeLevelScalar(level, NULL);
}

// =========================== 通用绘图函数 ===========================
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

void DrawRotatedText(HDC hdc, int x, int y, int w, int h, const char* text, int fontSize, double angle) {
	HFONT font = CreateFontA(fontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
							 DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
							 DEFAULT_QUALITY, DEFAULT_PITCH, "Microsoft YaHei");
	HFONT oldFont = (HFONT)SelectObject(hdc, font);
	SetTextColor(hdc, COLOR_TEXT);
	SetBkMode(hdc, TRANSPARENT);
	
	int centerX = x + w / 2;
	int centerY = y + h / 2;
	
	XFORM xform;
	xform.eM11 = (FLOAT)cos(angle);
	xform.eM12 = (FLOAT)sin(angle);
	xform.eM21 = (FLOAT)-sin(angle);
	xform.eM22 = (FLOAT)cos(angle);
	xform.eDx = (FLOAT)centerX - (FLOAT)centerX * xform.eM11 - (FLOAT)centerY * xform.eM21;
	xform.eDy = (FLOAT)centerY - (FLOAT)centerX * xform.eM12 - (FLOAT)centerY * xform.eM22;
	
	SetGraphicsMode(hdc, GM_ADVANCED);
	SetWorldTransform(hdc, &xform);
	
	RECT rect = {x, y, x + w, y + h};
	DrawTextA(hdc, text, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
	
	ModifyWorldTransform(hdc, NULL, MWT_IDENTITY);
	SelectObject(hdc, oldFont);
	DeleteObject(font);
}

// =========================== 辅助函数 ===========================
std::string ToLower(const std::string& str) {
	std::string result = str;
	std::transform(result.begin(), result.end(), result.begin(), ::tolower);
	return result;
}

bool IsTitleMatch(const char* title, const std::vector<std::string>& keywords) {
	std::string lowerTitle = ToLower(title);
	for (const auto& keyword : keywords) {
		if (lowerTitle.find(ToLower(keyword)) == std::string::npos) {
			return false;
		}
	}
	return true;
}

HWND FindPptWindow() {
	HWND result = NULL;
	EnumWindows([](HWND hwnd, LPARAM lparam) -> BOOL {
		if (IsWindowVisible(hwnd)) {
			char title[256];
			GetWindowTextA(hwnd, title, 256);
			
			std::vector<std::string> pptKeywords = {"powerpoint", "幻灯片放映"};
			std::vector<std::string> wpsKeywords = {"wps", "幻灯片放映"};
			
			if (IsTitleMatch(title, pptKeywords) || IsTitleMatch(title, wpsKeywords)) {
				*(HWND*)lparam = hwnd;
				return FALSE;
			}
		}
		return TRUE;
	}, (LPARAM)&result);
	return result;
}

void LoadConfig() {
	std::ifstream configFile("ppttool.config");
	if (configFile.is_open()) {
		std::string line;
		while (std::getline(configFile, line)) {
			line.erase(std::remove_if(line.begin(), line.end(), ::isspace), line.end());
			if (line.empty() || line[0] == '#') continue;
			
			size_t eqPos = line.find('=');
			if (eqPos != std::string::npos) {
				std::string key = line.substr(0, eqPos);
				std::string value = line.substr(eqPos + 1);
				
				if (key == "enableLeftNav") {
					g_config.enableLeftNav = (value == "1" || value == "true" || value == "True");
				} else if (key == "enableRightNav") {
					g_config.enableRightNav = (value == "1" || value == "true" || value == "True");
				} else if (key == "enableToolbar") {
					g_config.enableToolbar = (value == "1" || value == "true" || value == "True");
				} else if (key == "enableLeftBottomNav") {
					g_config.enableLeftBottomNav = (value == "1" || value == "true" || value == "True");
				} else if (key == "enableRightBottomNav") {
					g_config.enableRightBottomNav = (value == "1" || value == "true" || value == "True");
				}
			}
		}
		configFile.close();
	} else {
		std::ofstream outFile("ppttool.config");
		if (outFile.is_open()) {
			outFile << "# PPT 触屏工具配置文件\n";
			outFile << "# 1=启用，0=禁用\n\n";
			outFile << "enableLeftNav=1\n";
			outFile << "enableRightNav=1\n";
			outFile << "enableLeftBottomNav=1\n";
			outFile << "enableRightBottomNav=1\n";
			outFile << "enableToolbar=1\n";
			outFile.close();
		}
	}
}

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

// =========================== 音量弹出窗口过程（完全自绘滑块）===========================
LRESULT CALLBACK VolumePopupWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	static HWND hStatic = NULL;
	static HFONT hFont = NULL;
	
	switch (msg) {
		case WM_CREATE: {
		// 设置分层窗口实现透明
		SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_LAYERED);
		SetLayeredWindowAttributes(hwnd, COLOR_MAGIC, 127, LWA_COLORKEY | LWA_ALPHA);
		
		// 获取初始音量
		g_currentVolume = GetSystemVolume();
		
		// [修改点 1] 修复音量数字显示遮挡上边框
		// 原代码 y=2 太靠上，改为 y=4，并且高度稍微留余量
		hStatic = CreateWindowEx(0, "STATIC", NULL,
								 WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
								 155, 4, 35, 54,  // Y 从 2 改为 4，高度微调以适应内部垂直居中
								 hwnd, NULL, GetModuleHandle(NULL), NULL);
		
		// 设置字体
		hFont = CreateFontA(16, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
							DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
							DEFAULT_QUALITY, DEFAULT_PITCH, "Microsoft YaHei");
		SendMessage(hStatic, WM_SETFONT, (WPARAM)hFont, TRUE);
		
		char buf[8];
		sprintf(buf, "%d", g_currentVolume);
		SetWindowTextA(hStatic, buf);
		break;
	}
		
		case WM_SIZE: {
			// 动态计算滑块位置（上下居中）
			RECT rc;
			GetClientRect(hwnd, &rc);
			int clientHeight = rc.bottom;
			
			// 轨道垂直居中
			g_sliderY = (clientHeight - g_trackHeight) / 2;
			
			// 按钮位置
			g_thumbX = (int)((float)g_currentVolume / 100.0f * g_sliderWidth);
			if (g_thumbX < 0) g_thumbX = 0;
			if (g_thumbX > g_sliderWidth) g_thumbX = g_sliderWidth;
			
			// 更新数字显示位置 (保持 Y=4 的偏移，高度填满剩余空间)
			if (hStatic) {
				SetWindowPos(hStatic, NULL, 155, 4, 35, clientHeight - 8, SWP_NOZORDER);
			}
			break;
		}
		
		case WM_MOUSEMOVE: {
			if (g_dragging) {
				// 获取鼠标坐标并转换为客户区坐标
				POINT pt = { LOWORD(lParam), HIWORD(lParam) };
				ClientToScreen(hwnd, &pt);
				ScreenToClient(hwnd, &pt);
				int x = pt.x;
				
				// 计算滑块位置（边界安全处理）
				int newX = x - g_sliderX;
				
				// 严格边界检查
				if (newX < 0) {
					newX = 0;
				} else if (newX > g_sliderWidth) {
					newX = g_sliderWidth;
				}
				
				// 安全计算音量
				int newVolume = (newX * 100) / g_sliderWidth;
				newVolume = std::max(0, std::min(100, newVolume));
				
				if (newVolume != g_currentVolume) {
					g_currentVolume = newVolume;
					g_thumbX = newX;
					SetSystemVolume(g_currentVolume);
					
					char buf[8];
					sprintf(buf, "%d", g_currentVolume);
					SetWindowTextA(hStatic, buf);
					InvalidateRect(hwnd, NULL, TRUE);
				}
			}
			break;
		}
		
		case WM_LBUTTONDOWN: {
			POINT pt = { LOWORD(lParam), HIWORD(lParam) };
			ClientToScreen(hwnd, &pt);
			ScreenToClient(hwnd, &pt);
			int x = pt.x;
			int y = pt.y;
			
			int thumbCenterX = g_sliderX + g_thumbX;
			// 扩大点击区域以适应更窄的滑块
			int hitLeft = thumbCenterX - g_thumbWidth - 5; 
			int hitRight = thumbCenterX + g_thumbWidth + 5;
			int hitTop = g_sliderY - 10;
			int hitBottom = g_sliderY + g_trackHeight + 10;
			
			if ((x >= hitLeft && x <= hitRight && y >= hitTop && y <= hitBottom) ||
				(x >= g_sliderX && x <= g_sliderX + g_sliderWidth && y >= g_sliderY && y <= g_sliderY + g_trackHeight)) {
				g_dragging = true;
				SetCapture(hwnd);
				
				int newX = x - g_sliderX;
				if (newX < 0) newX = 0;
				if (newX > g_sliderWidth) newX = g_sliderWidth;
				
				int newVolume = (newX * 100) / g_sliderWidth;
				newVolume = std::max(0, std::min(100, newVolume));
				
				if (newVolume != g_currentVolume) {
					g_currentVolume = newVolume;
					g_thumbX = newX;
					SetSystemVolume(g_currentVolume);
					
					char buf[8];
					sprintf(buf, "%d", g_currentVolume);
					SetWindowTextA(hStatic, buf);
					InvalidateRect(hwnd, NULL, TRUE);
				}
			}
			break;
		}
		
		case WM_LBUTTONUP: {
			if (g_dragging) {
				g_dragging = false;
				ReleaseCapture();
			}
			break;
		}
		
		case WM_CAPTURECHANGED: {
			if (g_dragging) {
				g_dragging = false;
			}
			break;
		}
		
		case WM_ACTIVATE: {
			if (LOWORD(wParam) == WA_INACTIVE) {
				DestroyWindow(hwnd);
			}
			break;
		}
		
		case WM_CLOSE: {
			DestroyWindow(hwnd);
			break;
		}
		
		case WM_DESTROY: {
			if (hFont) DeleteObject(hFont);
			g_volumeWnd = NULL;
			g_dragging = false;
			break;
		}
		
	case WM_ERASEBKGND:
		return 1;
		
		case WM_PAINT: {
			PAINTSTRUCT ps;
			HDC hdc = BeginPaint(hwnd, &ps);
			
			RECT rc;
			GetClientRect(hwnd, &rc);
			int w = rc.right;
			int h = rc.bottom;
			
			// 使用魔术色填充背景
			HBRUSH magicBrush = CreateSolidBrush(COLOR_MAGIC);
			FillRect(hdc, &rc, magicBrush);
			DeleteObject(magicBrush);
			
			// [修改点 2] 修复边框被遮挡问题
			// 策略：绘制时严格控制在客户区内，留出 1px 给边框，不再向外扩展
			HBRUSH bgBrush = CreateSolidBrush(COLOR_BG);
			HPEN penBorder = CreatePen(PS_SOLID, 1, COLOR_BORDER);
			
			HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, bgBrush);
			HPEN oldPen = (HPEN)SelectObject(hdc, penBorder);
			
			// 绘制圆角矩形：从 (0,0) 到 (w-1, h-1)，确保边框完全在窗口内
			// 之前代码画到 w+2 导致右侧和下侧被裁剪
			RoundRect(hdc, 0, 0, w - 1, h - 1, 14, 14);
			
			SelectObject(hdc, oldBrush);
			SelectObject(hdc, oldPen);
			DeleteObject(bgBrush);
			DeleteObject(penBorder);
			
			// 绘制滑块背景
			HPEN nullPen = CreatePen(PS_NULL, 0, 0);
			HPEN oldPen2 = (HPEN)SelectObject(hdc, nullPen);
			
			HBRUSH sliderBgBrush = CreateSolidBrush(COLOR_SLIDER_BG);
			HBRUSH oldBrush2 = (HBRUSH)SelectObject(hdc, sliderBgBrush);
			RoundRect(hdc, g_sliderX, g_sliderY, g_sliderX + g_sliderWidth, g_sliderY + g_trackHeight, 4, 4);
			
			HBRUSH sliderFillBrush = CreateSolidBrush(COLOR_SLIDER_FILL);
			SelectObject(hdc, sliderFillBrush);
			RoundRect(hdc, g_sliderX, g_sliderY, g_sliderX + g_thumbX, g_sliderY + g_trackHeight, 4, 4);
			
			HBRUSH thumbBrush = CreateSolidBrush(COLOR_SLIDER_THUMB);
			SelectObject(hdc, thumbBrush);
			
			int thumbLeft = g_sliderX + g_thumbX - g_thumbWidth / 2;
			int thumbTop = (h - g_thumbHeight) / 2;
			int thumbRight = thumbLeft + g_thumbWidth;
			int thumbBottom = thumbTop + g_thumbHeight;
			
			if (thumbLeft < g_sliderX) thumbLeft = g_sliderX;
			if (thumbRight > g_sliderX + g_sliderWidth) thumbRight = g_sliderX + g_sliderWidth;
			
			RoundRect(hdc, thumbLeft, thumbTop, thumbRight, thumbBottom, 4, 4);
			
			SelectObject(hdc, oldBrush2);
			SelectObject(hdc, oldPen2);
			DeleteObject(nullPen);
			DeleteObject(sliderBgBrush);
			DeleteObject(sliderFillBrush);
			DeleteObject(thumbBrush);
			
			EndPaint(hwnd, &ps);
			return 0;
		}
		
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}

// 创建音量弹出窗口
void ShowVolumePopup(HWND parentWnd, int buttonIndex) {
	if (g_volumeWnd && IsWindow(g_volumeWnd)) {
		DestroyWindow(g_volumeWnd);
		g_volumeWnd = NULL;
		return;
	}
	
	POINT pt;
	RECT btnRect = g_buttonRects[buttonIndex];
	pt.x = btnRect.left;
	pt.y = btnRect.top;
	ClientToScreen(parentWnd, &pt);
	
	int popupWidth = 200;
	int popupHeight = 62;
	int x = pt.x + (btnRect.right - btnRect.left - popupWidth) / 2;
	int y = pt.y - popupHeight - 3;
	
	int screenWidth = GetSystemMetrics(SM_CXSCREEN);
	int screenHeight = GetSystemMetrics(SM_CYSCREEN);
	if (x + popupWidth > screenWidth) x = screenWidth - popupWidth - 2;
	if (x < 2) x = 2;
	if (y < 2) {
		y = pt.y + (btnRect.bottom - pt.y) + 3;
	}
	
	static bool classRegistered = false;
	if (!classRegistered) {
		WNDCLASSEXA wc = {0};
		wc.cbSize = sizeof(WNDCLASSEXA);
		wc.lpfnWndProc = VolumePopupWndProc;
		wc.hInstance = GetModuleHandle(NULL);
		wc.hCursor = LoadCursor(NULL, IDC_ARROW);
		wc.hbrBackground = NULL;
		wc.lpszClassName = "VolumePopupClass";
		RegisterClassExA(&wc);
		classRegistered = true;
	}
	
	// [修改点 2] 创建窗口时使用完整尺寸，不再减去 2px
	// 之前的减法导致客户区变小，而绘图逻辑没变，导致边缘被切
	g_volumeWnd = CreateWindowExA(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
								  "VolumePopupClass", "",
								  WS_POPUP,
								  x, y, popupWidth, popupHeight,
								  parentWnd, NULL, GetModuleHandle(NULL), NULL);
	
	// [修改点 2] 区域掩码也使用完整尺寸
	HRGN hRgn = CreateRoundRectRgn(0, 0, popupWidth, popupHeight, 14, 14);
	SetWindowRgn(g_volumeWnd, hRgn, TRUE);
	
	ShowWindow(g_volumeWnd, SW_SHOW);
	UpdateWindow(g_volumeWnd);
}

// =========================== 导航窗口过程 ===========================
LRESULT CALLBACK VerticalNavWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, bool isLeft) {
	static int btnHeight = 60;
	bool* prevPressed = isLeft ? &g_leftPrevPressed : &g_rightPrevPressed;
	bool* nextPressed = isLeft ? &g_leftNextPressed : &g_rightNextPressed;
	
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
			
			COLORREF prevColor = *prevPressed ? COLOR_ACTIVE : COLOR_BG;
			DrawRoundedRect(hdc, 1, 1, w - 2, btnHeight - 1, 7, prevColor);
			DrawText(hdc, 1, 1, w - 2, btnHeight - 1, "△", 40);
			
			HPEN pen = CreatePen(PS_SOLID, 1, COLOR_SEP);
			HPEN oldPen = (HPEN)SelectObject(hdc, pen);
			MoveToEx(hdc, 1, btnHeight, NULL);
			LineTo(hdc, w - 1, btnHeight);
			SelectObject(hdc, oldPen);
			DeleteObject(pen);
			
			COLORREF nextColor = *nextPressed ? COLOR_ACTIVE : COLOR_BG;
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
				*prevPressed = true;
				InvalidateRect(hwnd, NULL, TRUE);
				SendKey(VK_PRIOR);
			} else {
				*nextPressed = true;
				InvalidateRect(hwnd, NULL, TRUE);
				SendKey(VK_NEXT);
			}
			break;
		}
		
		case WM_LBUTTONUP: {
			*prevPressed = false;
			*nextPressed = false;
			InvalidateRect(hwnd, NULL, TRUE);
			break;
		}
		
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}

LRESULT CALLBACK HorizontalNavWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, bool isLeft) {
	static int btnWidth = 60;
	bool* prevPressed = isLeft ? &g_leftBottomPrevPressed : &g_rightBottomPrevPressed;
	bool* nextPressed = isLeft ? &g_leftBottomNextPressed : &g_rightBottomNextPressed;
	
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
			
			COLORREF prevColor = *prevPressed ? COLOR_ACTIVE : COLOR_BG;
			DrawRoundedRect(hdc, 1, 1, btnWidth - 1, h - 2, 7, prevColor);
			DrawRotatedText(hdc, 1, 1, btnWidth - 1, h - 2, "△", 40, -M_PI / 2);
			
			HPEN pen = CreatePen(PS_SOLID, 1, COLOR_SEP);
			HPEN oldPen = (HPEN)SelectObject(hdc, pen);
			MoveToEx(hdc, btnWidth, 1, NULL);
			LineTo(hdc, btnWidth, h - 1);
			SelectObject(hdc, oldPen);
			DeleteObject(pen);
			
			COLORREF nextColor = *nextPressed ? COLOR_ACTIVE : COLOR_BG;
			DrawRoundedRect(hdc, btnWidth + 1, 1, btnWidth - 1, h - 2, 7, nextColor);
			DrawRotatedText(hdc, btnWidth + 1, 1, btnWidth - 1, h - 2, "▽", 40, -M_PI / 2);
			EndPaint(hwnd, &ps);
			break;
		}
		
		case WM_LBUTTONDOWN: {
			int x = LOWORD(lParam);
			if (x < btnWidth) {
				*prevPressed = true;
				InvalidateRect(hwnd, NULL, TRUE);
				SendKey(VK_PRIOR);
			} else {
				*nextPressed = true;
				InvalidateRect(hwnd, NULL, TRUE);
				SendKey(VK_NEXT);
			}
			break;
		}
		
		case WM_LBUTTONUP: {
			*prevPressed = false;
			*nextPressed = false;
			InvalidateRect(hwnd, NULL, TRUE);
			break;
		}
		
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}

LRESULT CALLBACK LeftNavWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	return VerticalNavWndProc(hwnd, msg, wParam, lParam, true);
}

LRESULT CALLBACK RightNavWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	return VerticalNavWndProc(hwnd, msg, wParam, lParam, false);
}

LRESULT CALLBACK LeftBottomNavWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	return HorizontalNavWndProc(hwnd, msg, wParam, lParam, true);
}

LRESULT CALLBACK RightBottomNavWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	return HorizontalNavWndProc(hwnd, msg, wParam, lParam, false);
}

// =========================== 工具栏窗口过程 ===========================
LRESULT CALLBACK HorizontalWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_CREATE: {
		SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_LAYERED);
		SetLayeredWindowAttributes(hwnd, COLOR_MAGIC, 127, LWA_COLORKEY | LWA_ALPHA);
		
		g_buttonsPressed.resize(6, false);
		g_buttonRects.resize(6);
		g_buttonTexts = {"鼠标", "画笔", "橡皮", "清空", "音量", "结束放映"};
		break;
	}
		
	case WM_ERASEBKGND:
		return 1;
		
		case WM_SIZE: {
			RECT rect;
			GetClientRect(hwnd, &rect);
			int w = rect.right;
			int h = rect.bottom;
			int btnWidth = (w - 2 - 5) / 6;
			
			for (int i = 0; i < 6; i++) {
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
			
			for (int i = 0; i < 6; i++) {
				COLORREF btnColor = g_buttonsPressed[i] ? COLOR_ACTIVE : COLOR_BG;
				DrawRoundedRect(hdc, g_buttonRects[i].left, g_buttonRects[i].top,
								g_buttonRects[i].right - g_buttonRects[i].left,
								g_buttonRects[i].bottom - g_buttonRects[i].top, 7, btnColor);
				DrawText(hdc, g_buttonRects[i].left, g_buttonRects[i].top,
						 g_buttonRects[i].right - g_buttonRects[i].left,
						 g_buttonRects[i].bottom - g_buttonRects[i].top,
						 g_buttonTexts[i].c_str(), 20);
				
				if (i < 5) {
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
			
			if (g_volumeWnd && IsWindow(g_volumeWnd)) {
				DestroyWindow(g_volumeWnd);
				g_volumeWnd = NULL;
			}
			
			for (int i = 0; i < 6; i++) {
				if (x >= g_buttonRects[i].left && x <= g_buttonRects[i].right &&
					y >= g_buttonRects[i].top && y <= g_buttonRects[i].bottom) {
					g_buttonsPressed[i] = true;
					InvalidateRect(hwnd, NULL, TRUE);
					
					if (i == 4) {
						ShowVolumePopup(hwnd, i);
					} else {
						switch(i) {
						case 0:
							SendCtrlKey(VK_A);
							break;
						case 1:
							SendCtrlKey(VK_A);
							SendCtrlKey(VK_P);
							break;
						case 2:
							SendCtrlKey(VK_E);
							break;
						case 3:
							SendKey(VK_E);
							break;
						case 5:
							SendKey(VK_ESCAPE);
							SendKey(VK_ESCAPE);
							break;
						}
					}
					break;
				}
			}
			break;
		}
		
		case WM_LBUTTONUP: {
			for (int i = 0; i < 6; i++) {
				g_buttonsPressed[i] = false;
			}
			InvalidateRect(hwnd, NULL, TRUE);
			break;
		}
		
		case WM_KEYDOWN: {
			switch(wParam) {
			case VK_UP: case VK_LEFT:
				SendKey(VK_PRIOR);
				return 0;
			case VK_DOWN: case VK_RIGHT:
				SendKey(VK_NEXT);
				return 0;
			case 'M': case 'm':
				SendCtrlKey(VK_A);
				return 0;
			case 'P': case 'p':
				SendCtrlKey(VK_A);
				SendCtrlKey(VK_P);
				return 0;
			case 'E': case 'e':
				SendCtrlKey(VK_E);
				return 0;
			case 'C': case 'c':
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

// =========================== 工具函数 ===========================
HWND CreateToolWindow(HINSTANCE hInst, const char* className, WNDPROC wndProc, int width, int height, int x, int y) {
	WNDCLASSEXA wc = {0};
	wc.cbSize = sizeof(WNDCLASSEXA);
	wc.lpfnWndProc = wndProc;
	wc.hInstance = hInst;
	wc.hCursor = LoadCursor(NULL, IDC_HAND);
	wc.hbrBackground = (HBRUSH)GetStockObject(NULL_BRUSH);
	wc.lpszClassName = className;
	RegisterClassExA(&wc);
	
	HWND hwnd = CreateWindowExA(WS_EX_NOACTIVATE | WS_EX_TOPMOST,
								className, "", WS_POPUP,
								x, y, width + 2, height,
								NULL, NULL, hInst, NULL);
	
	HMENU menu = GetSystemMenu(hwnd, FALSE);
	if (menu) RemoveMenu(menu, SC_CLOSE, MF_BYCOMMAND);
	
	ShowWindow(hwnd, SW_HIDE);
	return hwnd;
}

void UpdateWindowsVisibility() {
	bool pptActive = (g_pptHwnd != NULL && IsWindow(g_pptHwnd));
	
	if (pptActive) {
		if (g_config.enableLeftNav && g_leftWnd) ShowWindow(g_leftWnd, SW_SHOW);
		if (g_config.enableRightNav && g_rightWnd) ShowWindow(g_rightWnd, SW_SHOW);
		if (g_config.enableLeftBottomNav && g_leftBottomWnd) ShowWindow(g_leftBottomWnd, SW_SHOW);
		if (g_config.enableRightBottomNav && g_rightBottomWnd) ShowWindow(g_rightBottomWnd, SW_SHOW);
		if (g_config.enableToolbar && g_horizontalWnd) ShowWindow(g_horizontalWnd, SW_SHOW);
	} else {
		if (g_leftWnd) ShowWindow(g_leftWnd, SW_HIDE);
		if (g_rightWnd) ShowWindow(g_rightWnd, SW_HIDE);
		if (g_leftBottomWnd) ShowWindow(g_leftBottomWnd, SW_HIDE);
		if (g_rightBottomWnd) ShowWindow(g_rightBottomWnd, SW_HIDE);
		if (g_horizontalWnd) ShowWindow(g_horizontalWnd, SW_HIDE);
		if (g_volumeWnd && IsWindow(g_volumeWnd)) DestroyWindow(g_volumeWnd);
	}
}

DWORD WINAPI MonitorThread(LPVOID lpParam) {
	HWND lastPptHwnd = NULL;
	
	while (true) {
		HWND pptHwnd = FindPptWindow();
		
		if (pptHwnd != lastPptHwnd) {
			lastPptHwnd = pptHwnd;
			g_pptHwnd = pptHwnd;
			UpdateWindowsVisibility();
		}
		
		Sleep(300);
	}
	return 0;
}

// =========================== 程序入口 ===========================
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpCmdLine, int nCmdShow) {
	INITCOMMONCONTROLSEX icc;
	icc.dwSize = sizeof(INITCOMMONCONTROLSEX);
	icc.dwICC = ICC_BAR_CLASSES;
	InitCommonControlsEx(&icc);
	
	LoadConfig();
	
	if (!InitVolumeControl()) {
		MessageBoxA(NULL, "警告：无法初始化音量控制，音量按钮可能无效。", "PPT 工具", MB_OK | MB_ICONWARNING);
	}
	
	int screenWidth = GetSystemMetrics(SM_CXSCREEN);
	int screenHeight = GetSystemMetrics(SM_CYSCREEN);
	
	if (g_config.enableLeftNav) {
		g_leftWnd = CreateToolWindow(hInst, "LeftNavWindow", LeftNavWndProc, 60, 121,
									 2, (screenHeight - 121) / 2);
	}
	
	if (g_config.enableRightNav) {
		g_rightWnd = CreateToolWindow(hInst, "RightNavWindow", RightNavWndProc, 60, 121,
									  screenWidth - 62, (screenHeight - 121) / 2);
	}
	
	if (g_config.enableLeftBottomNav) {
		g_leftBottomWnd = CreateToolWindow(hInst, "LeftBottomNavWindow", LeftBottomNavWndProc, 121, 60,
										   2, screenHeight - 62);
	}
	
	if (g_config.enableRightBottomNav) {
		g_rightBottomWnd = CreateToolWindow(hInst, "RightBottomNavWindow", RightBottomNavWndProc, 121, 60,
											screenWidth - 123, screenHeight - 62);
	}
	
	if (g_config.enableToolbar) {
		g_horizontalWnd = CreateToolWindow(hInst, "HorizontalWindow", HorizontalWndProc, 604, 45,
										   (screenWidth - 606) / 2, screenHeight - 45);
	}
	
	CreateThread(NULL, 0, MonitorThread, NULL, 0, NULL);
	
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	
	UninitVolumeControl();
	return 0;
}
