# TrayIcon

一个轻量级的 Windows 托盘程序，可从 CSV 文件动态生成菜单。

## 功能特点

- 在系统托盘显示图标和分层菜单
- 自动加载应用程序目录（含子目录）中的 `.csv` 文件
- 支持以下操作：
  - 复制内容到剪贴板
  - 打开路径或网址
  - 向活动窗口发送按键
- 支持菜单热重载，无需重启程序

## CSV 格式

每行包含 4 个字段：

MainMenu,SubMenu,DisplayItem,ContentItem


- `MainMenu`：主菜单名称
- `SubMenu`：子菜单名称（可为空）
- `DisplayItem`：菜单中显示的文字
- `ContentItem`：要复制、打开或发送的内容

## 使用方法

1. 将 `.csv` 文件放入可执行文件所在目录
2. 运行程序
3. 右键点击托盘图标以显示菜单
4. `Ctrl + 左键点击` 打开内容
5. `Ctrl + 右键点击` 发送内容为按键

## 系统要求

- .NET Framework 4.8 或更高版本
