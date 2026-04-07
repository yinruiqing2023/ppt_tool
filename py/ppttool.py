import tkinter as tk
from tkinter import messagebox
import ctypes
import time
import threading
import math

# Windows API 常量
VK_PRIOR = 0x21      # PageUp
VK_NEXT = 0x22       # PageDown
VK_ESCAPE = 0x1B
VK_CONTROL = 0x11
VK_M = 0x4D
VK_P = 0x50
VK_E = 0x45
VK_UP = 0x26
VK_DOWN = 0x28
VK_LEFT = 0x25
VK_RIGHT = 0x27

WS_EX_NOACTIVATE = 0x08000000
WS_EX_TRANSPARENT = 0x00000020
GWL_EXSTYLE = -20

user32 = ctypes.windll.user32

# 全局变量
ppt_hwnd = 0 

# 颜色定义
COLOR_MAGIC = '#FF00FF'      # 魔术色 (透明)
COLOR_BORDER = '#999999'     # 灰色边框
COLOR_BG = '#ffffff'         # 白色背景
COLOR_TEXT = '#000000'       # 黑色文字
COLOR_ACTIVE_BG = '#e6e6e6'  # 按下背景 (浅灰)
COLOR_SEP = '#999999'        # 分割线颜色

def find_ppt_window():
    """查找 PPT 放映窗口句柄"""
    hwnd_result = [0]
    def callback(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value
                if '幻灯片放映' in title or 'Slide Show' in title:
                    hwnd_result[0] = hwnd
                    return False
        return True
    user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)(callback), 0)
    return hwnd_result[0]

def activate_ppt():
    """激活 PPT 窗口并返回是否成功"""
    global ppt_hwnd
    if ppt_hwnd and user32.IsWindow(ppt_hwnd):
        hwnd = ppt_hwnd
    else:
        hwnd = find_ppt_window()
    
    if hwnd:
        ppt_hwnd = hwnd 
        user32.ShowWindow(hwnd, 5)      
        user32.SetForegroundWindow(hwnd)
        return True
    return False

def send_key(vk_code):
    """发送单键（先激活 PPT）"""
    if activate_ppt():
        user32.keybd_event(vk_code, 0, 0, 0)
        user32.keybd_event(vk_code, 0, 2, 0)

def send_ctrl_key(char_vk):
    """发送 Ctrl+键（无 Sleep，快速连续）"""
    if activate_ppt():
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        user32.keybd_event(char_vk, 0, 0, 0)
        user32.keybd_event(char_vk, 0, 2, 0)
        user32.keybd_event(VK_CONTROL, 0, 2, 0)

def set_window_no_focus(hwnd):
    """使窗口不获取焦点"""
    style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT)

def is_ppt_playing():
    """检测 PPT 是否正在放映"""
    result = [False]
    def callback(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value
                if '幻灯片放映' in title or 'Slide Show' in title:
                    result[0] = True
                    return False
        return True
    user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)(callback), 0)
    return result[0]

def draw_rounded_rect(canvas, x, y, width, height, radius, fill_color, outline_color=None, outline_width=0, tag=None):
    """绘制圆角矩形"""
    if radius <= 0:
        if outline_color and outline_width > 0:
            canvas.create_rectangle(x, y, x + width, y + height, fill=fill_color, 
                                    outline=outline_color, width=outline_width, tags=tag)
        else:
            canvas.create_rectangle(x, y, x + width, y + height, fill=fill_color, outline='', tags=tag)
        return
    
    r = min(radius, width / 2, height / 2)
    
    # 四个角圆弧
    canvas.create_arc(x, y, x + 2*r, y + 2*r, start=90, extent=90, fill=fill_color, outline='', tags=tag)
    canvas.create_arc(x + width - 2*r, y, x + width, y + 2*r, start=0, extent=90, fill=fill_color, outline='', tags=tag)
    canvas.create_arc(x + width - 2*r, y + height - 2*r, x + width, y + height, start=270, extent=90, fill=fill_color, outline='', tags=tag)
    canvas.create_arc(x, y + height - 2*r, x + 2*r, y + height, start=180, extent=90, fill=fill_color, outline='', tags=tag)
    
    # 边缘矩形
    canvas.create_rectangle(x + r, y, x + width - r, y + r, fill=fill_color, outline='', tags=tag)
    canvas.create_rectangle(x + r, y + height - r, x + width - r, y + height, fill=fill_color, outline='', tags=tag)
    canvas.create_rectangle(x, y + r, x + r, y + height - r, fill=fill_color, outline='', tags=tag)
    canvas.create_rectangle(x + width - r, y + r, x + width, y + height - r, fill=fill_color, outline='', tags=tag)
    
    # 中央填充
    canvas.create_rectangle(x + r, y + r, x + width - r, y + height - r, fill=fill_color, outline='', tags=tag)
    
    # 边框（如果需要）
    if outline_color and outline_width > 0:
        canvas.create_arc(x, y, x + 2*r, y + 2*r, start=90, extent=90, 
                         fill='', outline=outline_color, width=outline_width, tags=tag)
        canvas.create_arc(x + width - 2*r, y, x + width, y + 2*r, start=0, extent=90,
                         fill='', outline=outline_color, width=outline_width, tags=tag)
        canvas.create_arc(x + width - 2*r, y + height - 2*r, x + width, y + height, start=270, extent=90,
                         fill='', outline=outline_color, width=outline_width, tags=tag)
        canvas.create_arc(x, y + height - 2*r, x + 2*r, y + height, start=180, extent=90,
                         fill='', outline=outline_color, width=outline_width, tags=tag)
        
        canvas.create_line(x + r, y, x + width - r, y, fill=outline_color, width=outline_width, tags=tag)
        canvas.create_line(x + r, y + height, x + width - r, y + height, fill=outline_color, width=outline_width, tags=tag)
        canvas.create_line(x, y + r, x, y + height - r, fill=outline_color, width=outline_width, tags=tag)
        canvas.create_line(x + width, y + r, x + width, y + height - r, fill=outline_color, width=outline_width, tags=tag)


class VerticalWindow:
    """右侧翻页栏 (边框圆角8px，按钮圆角7px)"""
    def __init__(self, parent):
        self.parent = parent
        
        base_w = 60
        base_h = 121
        base_radius = 8      # 边框圆角半径 8px
        base_border = 1
        base_btn_radius = 7  # 按钮圆角半径 7px (比边框小1px)
        
        try:
            scale = 1.0
        except:
            scale = 1.0
        
        w = int(base_w * scale)
        h = int(base_h * scale)
        radius = int(base_radius * scale)
        border = int(base_border * scale)
        btn_radius = int(base_btn_radius * scale)
        
        self.window = tk.Toplevel()
        self.window.overrideredirect(True)
        self.window.attributes('-topmost', True)
        self.window.attributes('-alpha', 0.5)
        self.window.configure(bg=COLOR_MAGIC)
        self.window.attributes('-transparentcolor', COLOR_MAGIC)
        
        sw = self.window.winfo_screenwidth()
        sh = self.window.winfo_screenheight()
        self.window.geometry(f"{w}x{h}+{sw-w}+{(sh-h)//2}")
        
        self.canvas = tk.Canvas(self.window, bg=COLOR_MAGIC, highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        
        # 绘制整体边框背景
        draw_rounded_rect(self.canvas, 0, 0, w, h, radius, COLOR_BORDER, tag="border_bg")
        
        # 内层白色区域
        inner_w = w - 2 * border
        inner_h = h - 2 * border
        inner_r = radius - border
        draw_rounded_rect(self.canvas, border, border, inner_w, inner_h, inner_r, COLOR_BG, tag="inner_bg")
        
        btn_h = (inner_h - 1) // 2
        
        # 上一页按钮 (圆角半径 7px)
        self.prev_bg_tag = "prev_bg"
        draw_rounded_rect(self.canvas, border, border, inner_w, btn_h, btn_radius, COLOR_BG, tag=self.prev_bg_tag)
        self.canvas.create_text(border + inner_w/2, border + btn_h/2, text="△", 
                                font=('Segoe UI', int(28*scale)), fill=COLOR_TEXT, 
                                tags=("prev_text", "prev_btn"))
        self.canvas.addtag_withtag("prev_btn", self.prev_bg_tag)
        self.canvas.addtag_withtag("prev_btn", "prev_text")
        
        # 分割线
        line_y = border + btn_h
        self.canvas.create_rectangle(border, line_y, border + inner_w, line_y + 1, fill=COLOR_SEP, outline='')
        
        # 下一页按钮 (圆角半径 7px)
        self.next_bg_tag = "next_bg"
        draw_rounded_rect(self.canvas, border, line_y + 1, inner_w, btn_h, btn_radius, COLOR_BG, tag=self.next_bg_tag)
        self.canvas.create_text(border + inner_w/2, line_y + 1 + btn_h/2, text="▽",
                                font=('Segoe UI', int(28*scale)), fill=COLOR_TEXT,
                                tags=("next_text", "next_btn"))
        self.canvas.addtag_withtag("next_btn", self.next_bg_tag)
        self.canvas.addtag_withtag("next_btn", "next_text")
        
        self.canvas.bind('<Button-1>', self.on_click)
        self.canvas.bind('<ButtonRelease-1>', self.on_release)
        self.window.bind('<Key>', parent.on_key_press)
        
        self.window.after(100, self.set_no_focus)
        self.window.withdraw()
        self.active_tag = None
    
    def on_click(self, event):
        items = self.canvas.find_overlapping(event.x, event.y, event.x, event.y)
        for item in items:
            tags = self.canvas.gettags(item)
            if "prev_btn" in tags:
                self.active_tag = self.prev_bg_tag
                self.canvas.itemconfig(self.prev_bg_tag, fill=COLOR_ACTIVE_BG)
                self.parent.prev_page()
                break
            elif "next_btn" in tags:
                self.active_tag = self.next_bg_tag
                self.canvas.itemconfig(self.next_bg_tag, fill=COLOR_ACTIVE_BG)
                self.parent.next_page()
                break
    
    def on_release(self, event):
        if self.active_tag:
            self.canvas.itemconfig(self.active_tag, fill=COLOR_BG)
            self.active_tag = None
    
    def set_no_focus(self):
        try:
            hwnd = ctypes.c_int(self.window.winfo_id()).value
            set_window_no_focus(hwnd)
        except:
            pass
    
    def show(self):
        self.window.deiconify()
        self.set_no_focus()
    
    def hide(self):
        self.window.withdraw()


class HorizontalWindow:
    """底部工具栏 (边框圆角8px，按钮圆角7px)"""
    def __init__(self, parent):
        self.parent = parent
        
        num_buttons = 5
        base_btn_width = 100
        base_sep = 1
        base_h = 45
        base_radius = 8      # 边框圆角半径 8px
        base_border = 1
        base_btn_radius = 7  # 按钮圆角半径 7px (比边框小1px)
        
        try:
            scale = 1.0
        except:
            scale = 1.0
        
        btn_w = int(base_btn_width * scale)
        sep_w = int(base_sep * scale)
        h = int(base_h * scale)
        radius = int(base_radius * scale)
        border = int(base_border * scale)
        btn_radius = int(base_btn_radius * scale)
        
        content_w = (num_buttons * btn_w) + ((num_buttons - 1) * sep_w)
        w = content_w + 2 * border
        
        self.window = tk.Toplevel()
        self.window.overrideredirect(True)
        self.window.attributes('-topmost', True)
        self.window.attributes('-alpha', 0.5)
        self.window.configure(bg=COLOR_MAGIC)
        self.window.attributes('-transparentcolor', COLOR_MAGIC)
        
        sw = self.window.winfo_screenwidth()
        sh = self.window.winfo_screenheight()
        self.window.geometry(f"{w}x{h}+{(sw-w)//2}+{sh-h}")
        
        self.canvas = tk.Canvas(self.window, bg=COLOR_MAGIC, highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        
        # 整体边框背景
        draw_rounded_rect(self.canvas, 0, 0, w, h, radius, COLOR_BORDER, tag="toolbar_border")
        
        # 内层白色区域
        inner_w = w - 2 * border
        inner_h = h - 2 * border
        inner_r = radius - border
        draw_rounded_rect(self.canvas, border, border, inner_w, inner_h, inner_r, COLOR_BG, tag="toolbar_bg")
        
        buttons_config = [
            ("鼠标", parent.action_mouse),
            ("画笔", parent.action_pen),
            ("橡皮", parent.action_eraser),
            ("清空", parent.clear_ink),
            ("结束放映", parent.end_show)
        ]
        
        current_x = border
        self.button_bg_tags = []
        
        for idx, (text, cmd) in enumerate(buttons_config):
            bg_tag = f"btn_{idx}_bg"
            btn_tag = f"btn_{idx}"
            
            # 绘制圆角按钮背景 (圆角半径 7px)
            draw_rounded_rect(self.canvas, current_x, border, btn_w, inner_h, btn_radius, COLOR_BG, tag=bg_tag)
            self.canvas.addtag_withtag(btn_tag, bg_tag)
            
            # 绘制文字
            self.canvas.create_text(current_x + btn_w/2, border + inner_h/2, text=text,
                                    font=('微软雅黑', int(11*scale)), fill=COLOR_TEXT,
                                    tags=(f"btn_{idx}_txt", btn_tag))
            
            self.button_bg_tags.append(bg_tag)
            current_x += btn_w
            
            if idx < len(buttons_config) - 1:
                self.canvas.create_rectangle(current_x, border, current_x + sep_w, border + inner_h,
                                             fill=COLOR_SEP, outline='')
                current_x += sep_w
        
        self.canvas.bind('<Button-1>', self.on_click)
        self.canvas.bind('<ButtonRelease-1>', self.on_release)
        self.window.bind('<Key>', parent.on_key_press)
        
        self.window.after(100, self.set_no_focus)
        self.window.withdraw()
        self.active_tag = None
    
    def on_click(self, event):
        items = self.canvas.find_overlapping(event.x, event.y, event.x, event.y)
        for item in items:
            tags = self.canvas.gettags(item)
            for tag in tags:
                if tag.startswith("btn_") and tag != "btn_prev" and tag != "btn_next":
                    try:
                        idx = int(tag.split("_")[1])
                    except:
                        continue
                    
                    bg_tag = f"btn_{idx}_bg"
                    self.active_tag = bg_tag
                    self.canvas.itemconfig(bg_tag, fill=COLOR_ACTIVE_BG)
                    
                    actions = [
                        self.parent.action_mouse,
                        self.parent.action_pen,
                        self.parent.action_eraser,
                        self.parent.clear_ink,
                        self.parent.end_show
                    ]
                    if idx < len(actions):
                        actions[idx]()
                    return
    
    def on_release(self, event):
        if self.active_tag:
            self.canvas.itemconfig(self.active_tag, fill=COLOR_BG)
            self.active_tag = None
    
    def set_no_focus(self):
        try:
            hwnd = ctypes.c_int(self.window.winfo_id()).value
            set_window_no_focus(hwnd)
        except:
            pass
    
    def show(self):
        self.window.deiconify()
        self.set_no_focus()
    
    def hide(self):
        self.window.withdraw()


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        
        self.vertical = VerticalWindow(self)
        self.horizontal = HorizontalWindow(self)
        
        self.running = True
        self.last_state = False
        
        threading.Thread(target=self.monitor, daemon=True).start()
    
    def prev_page(self):
        send_key(VK_PRIOR)
    
    def next_page(self):
        send_key(VK_NEXT)
    
    def on_key_press(self, event):
        key_sym = event.keysym
        vk_map = {
            'Up': VK_UP,
            'Down': VK_DOWN,
            'Left': VK_LEFT,
            'Right': VK_RIGHT
        }
        if key_sym in vk_map:
            if activate_ppt():
                vk_code = vk_map[key_sym]
                user32.keybd_event(vk_code, 0, 0, 0)
                user32.keybd_event(vk_code, 0, 2, 0)
            return "break"
        return None
    
    def action_mouse(self):
        if activate_ppt():
            send_ctrl_key(VK_M)
            send_ctrl_key(VK_M)
    
    def action_pen(self):
        if activate_ppt():
            send_ctrl_key(VK_M)
            send_ctrl_key(VK_M)
            send_ctrl_key(VK_P)
    
    def action_eraser(self):
        if activate_ppt():
            send_ctrl_key(VK_E)
    
    def clear_ink(self):
        send_key(VK_E)
    
    def end_show(self):
        if activate_ppt():
            user32.keybd_event(VK_ESCAPE, 0, 0, 0)
            user32.keybd_event(VK_ESCAPE, 0, 2, 0)
            user32.keybd_event(VK_ESCAPE, 0, 0, 0)
            user32.keybd_event(VK_ESCAPE, 0, 2, 0)
    
    def monitor(self):
        while self.running:
            try:
                state = is_ppt_playing()
                if state != self.last_state:
                    self.last_state = state
                    if state:
                        self.root.after(0, self.show_windows)
                        global ppt_hwnd
                        ppt_hwnd = find_ppt_window()
                    else:
                        self.root.after(0, self.hide_windows)
                        ppt_hwnd = 0
                time.sleep(1)
            except Exception:
                time.sleep(1)
    
    def show_windows(self):
        self.vertical.show()
        self.horizontal.show()
    
    def hide_windows(self):
        self.vertical.hide()
        self.horizontal.hide()
    
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()
