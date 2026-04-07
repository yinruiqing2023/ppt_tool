# PPT 放映助手

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)
![C++](https://img.shields.io/badge/C%2B%2B-11-blue.svg)

## 项目简介

PPT 放映助手是一个专为触控教学大屏设计的轻量级工具，解决 Office COM/ROT（Running Object Table）失效时无法通过常规方式控制 PPT 放映的问题。

### 适用场景

- 教室/会议室触控大屏，无键盘操作
- 希沃PPT小工具等现有方案无法使用时
- Office COM 组件无法正常调用
- ROT 表中无法找到放映中的 PPT 实例
- 需要不依赖 Office API 的通用解决方案

### 核心特性(c++版本)

| 特性 | 说明 |
|------|------|
| 支持触控 | --- |
| 音量控制 | 内置滑块，可调节系统音量（特色功能） |
| 多位置导航 | 支持左中、右中、左下、右下四个位置，可单独开关 |
| 极致轻量 | 单文件 EXE，约 1MB，无外部依赖 |
| 低资源占用 | 内存约 2MB，CPU 接近 0% |
| 透明界面 | 半透明浮动窗口，不影响观看 |
| 可配置 | 通过配置文件启用/禁用各组件 |

---

## 版本对比

### C++ 版本（主分支 - 推荐使用）

| 项目 | 说明 |
|------|------|
| 状态 | 积极维护 |
| 文件大小 | 约 1MB（单 EXE，静态链接） |
| 内存占用 | 约 2MB |
| 依赖 | 无（纯 Win32 API） |
| 系统要求 | Windows 8 / 10 / 11（Win7 需 KB2999226 补丁） |
| 启动速度 | <10ms |
| 音量控制 | 滑块式系统音量调节 |
| 配置方式 | ppttool.config 文件 |

### Python 版本（legacy 分支 - 停止维护）

| 项目 | 说明 |
|------|------|
| 状态 | 仅作参考，不再维护 |
| 依赖 | Python 3.7+ / tkinter |
| 系统要求 | Windows 7+（源码运行）release仅提供win10+版|
| 音量控制 | 无 |
| 配置方式 | 硬编码 |

### 快捷键差异说明

两个版本的键盘快捷键实现不同（均为备用方案，日常使用推荐触摸屏按钮）：

| 功能 | Python 旧版 | C++ 新版 |
|------|-------------|----------|
| 鼠标模式 | Ctrl+M（两次） | Ctrl+A |
| 画笔模式 | Ctrl+M（两次）→ Ctrl+P | Ctrl+A → Ctrl+P |

---

## 快速开始

### 下载使用

1. 从 Releases 下载最新版 `ppttool_cpp.exe`
2. 双击运行（无需管理员权限）
3. 打开 PowerPoint/WPS 并开始放映
4. 工具自动出现在屏幕指定位置

### 配置文件

创建 `ppttool.config` 文件（与 EXE 同目录）：

```ini
enableLeftNav=1
enableRightNav=1
enableLeftBottomNav=1
enableRightBottomNav=1
enableToolbar=1
```

配置项说明：

| 配置项 | 说明 |
|--------|------|
| enableLeftNav | 左侧中部翻页栏 |
| enableRightNav | 右侧中部翻页栏 |
| enableLeftBottomNav | 左下角翻页栏 |
| enableRightBottomNav | 右下角翻页栏 |
| enableToolbar | 底部工具栏 |

---

## 界面说明

### 翻页栏（垂直）

位于屏幕左侧中部或右侧中部，尺寸 60x121px：
- 点击上箭头（△）触发上一页（PageUp）
- 点击下箭头（▽）触发下一页（PageDown）

### 翻页栏（水平）

位于屏幕左下角或右下角，尺寸 121x60px，适合角落单手操作。

### 音量滑块

点击「音量」按钮后弹出滑块，可拖动调节 Windows 系统音量，点击外部自动关闭。

---

## 按钮功能对照表

| 按钮 | 快捷键 | 功能 |
|------|--------|------|
| 鼠标 | Ctrl+A | 鼠标模式 |
| 画笔 | Ctrl+A → Ctrl+P | 画笔模式 |
| 橡皮 | Ctrl+E | 橡皮模式 |
| 清空 | E | 清空笔迹 |
| 音量 | - | 弹出音量滑块 |
| 结束放映 | ESC（两次） | 结束放映（画笔模式下需按两次） |
| 上一页 | ↑ / ← | 上一页 |
| 下一页 | ↓ / → | 下一页 |

---

## DPI 缩放说明

本程序未实现 DPI 自动缩放。若在高分辨率屏幕下按钮偏小或显示模糊，请右键 exe → 属性 → 兼容性 → 更改高 DPI 设置 → 勾选「替代高 DPI 缩放行为」→ 选择「系统（增强）」，测试在125％下完美显示。

---

## 编译指南

### C++ 版本

#### 系统要求

- Windows 8 / 10 / 11
- Windows 7 需安装 KB2999226 补丁（Universal C Runtime）
- 注：Release 包使用 GCC 11 编译，未在 Windows 7 无补丁环境测试，可自编译旧版GCC以兼容旧版Windows

#### MinGW (GCC 11 推荐)

```bash
g++ -mwindows -static -O2 -o ppttool.exe ppttool.cpp -lcomctl32 -lole32 -lgdi32 -luser32 -lwinmm -static-libgcc -static-libstdc++
```


#### MSVC (Visual Studio)

```bash
cl /O2 /MT ppttool.cpp user32.lib gdi32.lib ole32.lib uuid.lib comctl32.lib winmm.lib /link /SUBSYSTEM:WINDOWS
```

### Python 版本（legacy 分支）

```bash
# 运行
python ppt_assistant.py

# 打包为 EXE
pip install pyinstaller
pyinstaller --onefile --noconsole ppt_assistant.py
```

---


## 技术架构

### 核心原理

1. EnumWindows 枚举所有窗口
2. 匹配标题包含「幻灯片放映」或「Slide Show」
3. 获取窗口句柄后，通过 keybd_event 发送虚拟按键
4. 绕过 COM/ROT，直接控制 PPT 窗口

### 音量控制

使用 Windows Core Audio API 的 IAudioEndpointVolume 接口实现系统音量调节。

### 与 COM/ROT 方案对比

| 方案 | 原理 | 优点 | 缺点 |
|------|------|------|------|
| COM/ROT | 通过 GetActiveObject 获取 PPT 对象 | 功能完整 | Office COM 失效时不可用 |
| 本工具 | 模拟键盘按键控制窗口 | 通用性强，不依赖 Office | 无法读取 PPT 内部状态 |

---

## Star

如果这个项目对你有帮助，请 Star 支持一下～

---

## 免责声明

本工具仅供教学和个人使用。请勿用于任何商业用途或违反当地法律法规的场景。使用本工具产生的任何后果由使用者自行承担。

---

## 开源协议

MIT 协议
