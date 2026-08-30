<div align="center">
  <a href="https://github.com/Kuingsmile/word-GPT-Plus">
    <img src="./public/logo.svg" alt="Logo" height="100" >
  </a>

  <h2 align="center">Word GPT Plus</h2>
  <p align="center">
    将 AI 和 Agent 直接集成到 Microsoft Word
    <br />
    <a href="https://github.com/Kuingsmile/word-GPT-Plus/blob/master/LICENSE">
      <img src="https://img.shields.io/github/license/Kuingsmile/word-GPT-Plus?style=flat-square" alt="license" />
    </a>
    <a href="https://github.com/Kuingsmile/word-GPT-Plus/releases">
      <img src="https://img.shields.io/github/v/release/Kuingsmile/word-GPT-Plus?style=flat-square" alt="release" />
    </a>
    <a href="https://github.com/Kuingsmile/word-GPT-Plus/stargazers">
      <img src="https://img.shields.io/github/stars/Kuingsmile/word-GPT-Plus?style=flat-square" alt="stars" />
    </a>
    <br />
    <a href="#功能特点">功能特点</a> •
    <a href="#开始使用">开始使用</a> •
    <a href="#安装说明">安装说明</a> •
    <a href="#使用方法">使用方法</a>
  </p>
</div>

简体中文 | [English](https://github.com/gzsteven666/word-GPT-Plus/blob/master/README.md)

## 📋 简介

Word GPT Plus 是一款将 AI 和 Agent 无缝集成到 Microsoft
Word 中的插件，使您能够在文档中直接生成、翻译、总结和润色文本。增强您的写作流程，无需离开 Word 环境。

![Image](https://github.com/user-attachments/assets/e5b077ca-b8d4-4e28-97c7-b708524e1188)

## ✨ 功能特点

- **多平台支持**：
  - **OpenAI**：最新GPT系列（兼容 DeepSeek 和其他 OpenAI 兼容接口）
  - **Azure OpenAI**：完整的 Azure 集成，支持自定义部署名称
  - **Google Gemini**：Gemini 3 Pro/Flash、Gemini 2.5 Pro/Flash、AQA
  - **Ollama**：仅限本地部署时可用
  - **Groq**：Llama 3.3/4、Qwen3、Kimi-K2 等
  - 以上平台均支持自定义模型名称

- **智能 Agent 模式**：
  - **直接操作 Word 文档**：Agent 可以读取、写入和修改您的 Word 文档
  - **多种内置工具**：网络查询和访问、插入文本、格式化内容、创建表格、管理书签、搜索替换等
  - 多步骤推理，支持对话记忆
  - 实时流式响应
  - 思考过程可视化（可折叠显示）

- **双模式切换**：
  - **标准聊天**：快速问答和内容生成
  - **Agent 模式**：高级文档操作，可使用内置工具

- **快捷操作**：
  - 一键翻译（支持 40+ 种语言）
  - 文本润色和改进
  - 学术写作增强
  - 内容摘要生成
  - 语法检查和修正

- **自定义和灵活性**：
  - 为每个提供商添加自定义模型
  - 保存和管理自定义提示词
  - 按提供商调整温度和最大令牌数
  - 支持自定义 Base URL 和代理
  - 本地存储保护隐私
  - 多语言界面（English、简体中文）

- **高级格式化**：
  - **自动 Word 格式化**：插入聊天内容时自动应用适当的 Word 样式
  - Markdown 解析并转换为 Word 格式
  - 每条消息支持替换、追加或复制操作

- **现代化用户体验**：
  - 简洁的 Copilot 风格界面
  - 实时流式响应
  - 快速切换模型和提供商
  - 随时中止生成

## 🚀 开始使用

### 环境要求

#### 软件

- Microsoft Word 2016/2019 零售版、Word 2021 或 Microsoft 365
- [Edge WebView2 运行时](https://developer.microsoft.com/en-us/microsoft-edge/webview2/)
- Node.js 20+（仅用于自托管）

> **注意**：仅适用于 .docx 文件（不兼容旧版 .doc 格式）

#### API 访问

- **OpenAI**：从 [OpenAI Platform](https://platform.openai.com/account/api-keys) 获取 API 密钥
- **Azure OpenAI**：在 [Azure OpenAI Service](https://go.microsoft.com/fwlink/?linkid=2222006) 申请访问权限
- **Google Gemini**：从 [Google AI Studio](https://developers.generativeai.google/) 请求 API 访问
- **Groq**：从 [Groq Console](https://console.groq.com/keys) 获取 API 密钥

## 💻 安装说明

选择最适合的安装方式：

### 方式一：即刻使用（Windows / macOS）

_适合大多数用户，无需构建源码。应用由 [GitHub Pages](https://gzsteven666.github.io/word-GPT-Plus/)
托管，后续版本会在重新打开加载项时自动更新。_

#### Windows 一键安装

以管理员身份打开 PowerShell，运行：

```powershell
$script = "$env:TEMP\install-word-gpt.ps1"
Invoke-WebRequest https://raw.githubusercontent.com/gzsteven666/word-GPT-Plus/master/scripts/install-windows.ps1 -OutFile $script
powershell -ExecutionPolicy Bypass -File $script
```

保存文档并重启 Word，然后选择 **插入 → 获取加载项 → 共享文件夹 → GPT Plus Steven**。

#### macOS 一键安装（推荐）

适用于 Microsoft Word for Mac 桌面版。该脚本只会把加载项清单安装到当前用户目录，不需要 `sudo`，也不会上传或读取您的 API
Key。

1. 打开一次 Microsoft Word，然后保存并完全退出 Word（按 `Command + Q`）。
2. 打开“终端”，复制并执行：

   ```bash
   curl -fsSL https://raw.githubusercontent.com/gzsteven666/word-GPT-Plus/master/scripts/install-macos.sh -o /tmp/install-word-gpt.sh
   bash /tmp/install-word-gpt.sh
   ```

3. 重新打开 Word，并打开一个 `.docx` 文档。
4. 选择 **主页 → 加载项 → GPT Plus Steven**，打开插件。
5. 在插件的“设置”中填写您自己的 AI 服务商和 API Key。

以后更新版本时重新执行上面的命令即可。若 Word 正在运行，脚本会提示您保存文档、按 `Command + Q`
完全退出 Word，再重新打开。

> **看不到加载项时**：确认文件已经安装到
> `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`，然后完全退出并重启 Word。仍未刷新时，可执行
> `npx office-addin-cache clear` 后再次重启 Word。

#### macOS 手动安装

微软官方说明：[在 Mac 上旁加载 Office 加载项](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)。如果不想运行脚本，也可以手动安装：

1. 下载个人版
   [manifest.xml](https://raw.githubusercontent.com/gzsteven666/word-GPT-Plus/master/release/instant-use/manifest.xml)。
2. 在 Finder 中按 `Command + Shift + G`，输入并打开： `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
3. 如果 `wef` 目录不存在，请手动创建。
4. 将下载的文件复制到该目录，并命名为 `GPT-Plus-Steven.xml`。
5. 完全退出并重新打开 Word，打开 `.docx` 文档后选择 **主页 → 加载项 → GPT Plus Steven**。

#### 卸载 macOS 加载项

退出 Word 后，在“终端”执行：

```bash
rm -f "$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/GPT-Plus-Steven.xml"
```

如果脚本生成了备份文件且不再需要，也可以删除同目录下的 `GPT-Plus-Steven.xml.backup`。

#### 手动安装

下载个人版
[manifest.xml](https://github.com/gzsteven666/word-GPT-Plus/blob/master/release/instant-use/manifest.xml)，然后按照下方对应平台的[旁加载插件指南](#旁加载插件)操作。

> OpenAI 兼容接口仍需允许来自 Office WebView 的跨域请求；遇到 CORS 限制时请使用支持浏览器访问的网关或自托管方式。

### 方式二：自托管（高级）

_适合开发人员或需要私有后端的用户。_

<details>
<summary><strong>Docker 部署</strong></summary>

1. 拉取并运行 Docker 镜像：

   ```bash
   docker pull kuingsmile/word-gpt-plus
   docker run -d -p 3000:80 kuingsmile/word-gpt-plus
   ```

2. 下载 [manifest.xml](https://github.com/Kuingsmile/word-GPT-Plus/blob/master/release/self-hosted/manifest.xml)。
3. 编辑 `manifest.xml`：将所有 `http://localhost:3000` 替换为您的服务器地址。
4. 继续阅读下方的 [旁加载插件指南](#旁加载插件)。

</details>

<details>
<summary><strong>源码构建</strong></summary>

_需要 Node.js 20+_

1. 克隆并启动项目：

   ```bash
   git clone https://github.com/Kuingsmile/Word-GPT-Plus.git
   cd Word-GPT-Plus
   yarn
   yarn run build
   yarn run serve
   ```

2. 使用
   [自托管 manifest.xml](https://github.com/Kuingsmile/word-GPT-Plus/blob/master/release/self-hosted/manifest.xml)。
3. 继续阅读下方的 [旁加载插件指南](#旁加载插件)。

</details>

<details>
<summary><strong>部署到腾讯 EdgeOne</strong></summary>

[![使用 EdgeOne Pages 部署](https://cdnstatic.tencentcs.com/edgeone/pages/deploy.svg)](https://edgeone.ai/pages/new?repository-url=https%3A%2F%2Fgithub.com%2FKuingsmile%2FWord-GPT-Plus%2Ftree%2Fmaster&build-command=npm%20run%20build&output-directory=.%2Fdist&install-command=yarn%20install)

</details>

## 旁加载插件

为了开始使用 Word GPT Plus，你需要将插件旁加载到 Microsoft Word 中。

### Windows 手动旁加载

微软官方说明：[通过网络共享旁加载 Office 加载项](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)

1. 打开你保存了 `manifest.xml` 文件的文件夹，例如 `C:\Users\username\Documents\WordGPT`.
2. 右键点击文件夹打开菜单，选择 **属性**.
3. 在 **属性** 对话框中，选择 **共享** 选项卡，然后选择 **共享**.
   ![image](https://learn.microsoft.com/en-us/office/dev/add-ins/images/sideload-windows-properties-dialog.png)
4. 在 **网络访问** 对话框中，添加你自己和任何你想要共享的其他用户，选择 **共享**
   按钮，当你看到你的文件夹被共享的确认信息时，注意显示在文件夹名称后面的 **完整网络路径**.
   ![image](https://learn.microsoft.com/en-us/office/dev/add-ins/images/sideload-windows-network-access-dialog.png)
5. 在 Word 中打开一个新文档，选择 **文件** 选项卡，然后选择 **选项**.
6. 选择 **信任中心**，然后选择 **信任中心设置** 按钮.
7. 选择 **信任的目录**.
8. 在 **目录 URL** 框中，输入 **完整网络路径**，然后选择 **添加目录**.
9. 选择 **在菜单中显示** 复选框，然后选择 **确定**.
   ![image](https://learn.microsoft.com/en-us/office/dev/add-ins/images/sideload-windows-trust-center-dialog.png)
10. 关闭并重新启动 Word.
11. 点击**插入** -> **获取加载项** -> **共享目录**，选择 **Word GPT**.
12. 享受 Word GPT Plus 的强大功能吧！
    ![image](https://user-images.githubusercontent.com/96409857/234744280-9d9f13cf-536b-4fb5-adfa-cbec262d56a2.png)

## 📖 使用方法

### 快速开始

进入 Word GPT Plus 后，点击主页的`设置`图标，配置您首选的 AI 提供商和 API 密钥。

### 双模式

#### 聊天模式

使用标准聊天模式进行：

- 快速问答和信息查询
- 内容生成和头脑风暴
- 语言翻译
- 文本改进和润色

只需输入您的消息并按 Enter 键。使用快捷操作按钮执行常见任务。

#### Agent 模式

Agent 模式通过 25+ 种专用工具让 AI 直接访问您的 Word 文档：

**通用工具：**

- 网络查询和访问
- 数字计算
- 日期获取

**文档读取工具：**

- 获取选中文本或完整文档内容
- 搜索特定文本
- 获取文档属性和表格信息

**文档写入工具：**

- 在特定位置插入、替换或追加文本
- 创建带样式的格式化段落
- 插入表格和列表
- 添加书签和内容控件

**文档格式化工具：**

- 应用粗体、斜体、下划线和其他格式
- 更改字体名称和样式
- 清除格式
- 搜索和替换文本模式

**Agent 提示词示例：**

- "阅读整个文档并在开头创建摘要"
- "将所有节标题格式化为标题 2 并设为蓝色"
- "在第一个标题后插入解释主题的段落"

### 快捷操作

点击快捷操作按钮进行即时操作：

- **🌐 翻译**：将选中文本翻译为您偏好的语言
- **✨ 润色**：提高写作质量和清晰度
- **📚 学术**：增强学术写作
- **📝 总结**：创建简洁摘要
- **✔️ 语法**：检查和修正语法

### 自定义模型

对于每个 AI 提供商，您可以添加自定义模型：

1. 进入 设置 → 提供商 选项卡
2. 选择您的提供商
3. 输入自定义模型名称并点击添加
4. 模型将出现在模型下拉列表中

### 配置提示

- **温度（Temperature）**：较低（0.3-0.5）用于事实性任务，较高（0.7-1.0）用于创意任务
- **最大令牌数（Max Tokens）**：增加以获得更长响应，减少以获得简洁答案
- **自定义 Base URL**：用于 OpenAI 兼容服务，如 DeepSeek

## 🔒 隐私与安全

- **本地存储**：您的 API 密钥和自定义提示词存储在浏览器本地存储中（在 Word 插件环境内）。它们永远不会发送到我们的服务器。
- **直接连接**：插件直接与 AI 提供商（OpenAI、Azure 等）或您的本地 Ollama 实例通信。除非您使用自定义代理，否则没有中间服务器处理您的数据。

## 贡献

如果你希望贡献代码，请 fork 这个仓库并创建一个 pull request。

## License

MIT License

## 请给个 ⭐️ 吧

如果这个项目帮助到了你，请给个 ⭐️ 吧！
