# Zotero Linker

中文 | [English](README.en.md)

Zotero Linker 是一组面向 Windows 的 Zotero Office 插件，包含 Word 引文导航插件和 PowerPoint 引用插件。Word 插件让正文引用与参考文献双向跳转；PowerPoint 插件把 Zotero 的样式选择、文献选择、编号引用和参考文献列表带进幻灯片。

## 解决的 Gap

### Word：引用存在，但不便导航

Zotero 可以在 Word 中插入引用并生成参考文献，但正文中的 `[1]`、`[2-4]` 或 `[3, 5, 8]` 与文末条目之间通常缺少便捷的双向导航。长文档需要频繁滚动和查找，影响写作、阅读和审阅效率。

Word 插件保留 Zotero 原有引用域，为正文引用和参考文献条目建立双向链接，并提供格式修复和链接清理功能。

### PowerPoint：缺少 Zotero 原生引用工作流

Zotero 官方 Office 集成主要面向 Word，PowerPoint 中没有同等的样式选择、文献检索、多选引用、自动编号和参考文献刷新流程。用户通常只能手工输入 `[1]`、维护编号并复制参考文献；增删文献后容易出现编号错位、样式不一致和列表遗漏。

PowerPoint 插件通过 Zotero 本地引用协议打开 Zotero 的文档首选项和文献选择窗口，将单篇或多篇引用写入幻灯片，并生成和刷新参考文献页。它不替代 Zotero，而是补齐 Zotero 与演示文稿之间的集成空白。

## Word 插件

- 正文引用跳转到参考文献，并支持从参考文献返回正文引用。
- 支持 `[1]`、`[1,3,5]`、`[2-4]` 等数字引用形式。
- 结合 Zotero 域信息处理压缩引用中的可见和隐藏项目。
- 修复引用颜色、下划线和字号。
- 删除插件生成的链接和书签，同时保留 Zotero 引用域。
- 支持 Microsoft Office Word 和 WPS Word/Writer。

## PowerPoint 插件

- 首次使用时打开 Zotero 文档首选项，选择 CSL 引用样式。
- 调用 Zotero 文献选择窗口，支持单选和多选引用。
- 在当前幻灯片插入编号引用，例如 `[1]` 或 `[1–4]`。
- 新建或更新 `References` 幻灯片中的参考文献列表。
- 支持 `Document Preferences` 和 `Refresh`。
- 在演示文稿中保存 Zotero 文档和引用字段元数据，便于后续刷新。

## 兼容性

| 插件 | 支持环境 |
| --- | --- |
| Word | Windows、Microsoft Office Word、WPS Word/Writer、Zotero |
| PowerPoint | Windows、Microsoft Office PowerPoint、Zotero |
| WPS 演示 | 安装器会写入 WPS Presentation (`WPP`) 加载项白名单；实际加载取决于所用 WPS 版本对 VSTO 的兼容性 |

PowerPoint 插件要求 Zotero 正在运行，并允许本地应用通信。引用和参考文献以 PowerPoint 文本形状保存，而不是 Word 域。

## 安装

从 [GitHub Releases](https://github.com/Yccc1220/ZoteroLinker/releases/latest) 下载对应安装包：

```text
ZoteroLinkerSetup.exe       # Word / WPS Writer
ZoteroLinkerPptSetup.exe    # PowerPoint / WPS Presentation compatibility registration
```

以管理员身份运行安装包，安装完成后重新打开对应 Office 或 WPS 应用。

## PowerPoint 基本使用

1. 启动 Zotero，然后打开 PowerPoint 演示文稿。
2. 在 `Zotero Linker` 功能区点击 `Insert Citation`。
3. 首次使用时选择引用样式；随后在 Zotero 窗口中选择一篇或多篇文献。
4. 点击 `Add Bibliography` 生成参考文献页。
5. 文献或样式发生变化后，点击 `Refresh` 更新引用和参考文献。

## 仓库内容

- `Zotero-linker/`：Word VSTO 插件源码和安装器。
- `Zotero-linker-ppt/`：PowerPoint VSTO 插件源码和安装器。
- `release/`：Word、PowerPoint 安装包及 SHA-256 校验文件。
- `.github/workflows/release.yml`：GitHub Release 发布流程。
- `site/`：GitHub Pages 下载页面。
