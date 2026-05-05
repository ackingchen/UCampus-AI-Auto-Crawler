# UCampus AI Auto Crawler 使用说明

`UCampus AI Auto Crawler` 用于自动爬取网页课程中所有单元 `Further listening` 部分的题目内容，并生成 Word 文档。

脚本会自动遍历每个单元下的：

- `Conversation`
- `Passage`
- `Lectures`
- `Lectures` 中的 `Lecture 1 / Lecture 2` 等子标签

爬取完成后会生成：

- 每个分类单独的 `.docx` 文件
- 一份总汇总文档：`全部题目汇总.docx`

---

## 一、使用前准备

### 1. 安装 Node.js

请先安装 Node.js，建议安装 LTS 版本。

下载地址：

```text
https://nodejs.org/
```

安装完成后，打开 PowerShell，输入：

```powershell
node -v
npm -v
```

如果能看到版本号，说明安装成功。

### 2. 准备浏览器

本项目支持：

- Google Chrome
- Microsoft Edge

推荐使用 Chrome 或 Edge 的独立调试窗口运行，不要用平时正在使用的浏览器窗口。

---

## 二、安装项目依赖

进入项目目录：

```powershell
cd "<你的项目路径>\English4"
```

安装依赖：

```powershell
npm install
```

如果第一次使用 Playwright，也可以执行：

```powershell
npm run install-browser
```

---

## 三、启动可被脚本控制的浏览器

脚本不是直接打开网页，而是连接一个“远程调试模式”的浏览器。

### 如何获得临时浏览器配置目录

`--user-data-dir` 后面的 `<你的临时浏览器配置目录>` 不是浏览器自动提供的路径，而是你自己指定的一个临时文件夹。这个文件夹用于保存本次调试浏览器的登录状态、缓存和 Cookie，建议和日常使用的浏览器配置分开，避免影响平时的浏览器数据。

最简单的做法是在项目目录下新建一个临时目录，例如：

```powershell
New-Item -ItemType Directory -Force .\browser-profile
```

然后把命令里的 `<你的临时浏览器配置目录>` 替换为这个目录的完整路径。假设项目在桌面，可以写成：

```text
C:\Users\你的用户名\Desktop\UCampus AI Auto Crawler\browser-profile
```

也可以在 PowerShell 中进入项目目录后，用下面的命令查看它的完整路径：

```powershell
Resolve-Path .\browser-profile
```

例如，Chrome 启动命令可以改成：

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Users\你的用户名\Desktop\UCampus AI Auto Crawler\browser-profile"
```

注意：这个目录不要使用平时 Chrome/Edge 正在使用的用户目录，也不要直接填浏览器安装目录。

### 方式一：使用 Chrome

在 PowerShell 执行：

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="<你的临时浏览器配置目录>"
```

### 方式二：使用 Edge

在 PowerShell 执行：

```powershell
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="<你的临时浏览器配置目录>"
```

如果提示找不到浏览器，请把命令中的浏览器路径替换为你电脑上的实际安装路径。

Windows 常见浏览器路径示例：

```text
C:\Program Files\Google\Chrome\Application\chrome.exe
C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
```

---

## 四、登录课程网站并打开课程页面

浏览器启动后，会出现一个新的 Chrome/Edge 窗口。

请在这个新窗口中：

1. 打开课程网站
2. 登录账号
3. 进入课程页面
4. 确保右侧目录能看到各个单元，例如 `Unit 1`、`Unit 2`
5. 页面中能看到 `Further listening` 相关目录

不需要手动点完所有题目，脚本会自动点击。

---

## 五、开始批量爬取

回到 PowerShell，进入项目目录：

```powershell
cd "<你的项目路径>\UCampus AI Auto Crawler"
```

执行：

```powershell
npm run crawl
```

脚本会自动执行以下流程：

1. 连接 `http://127.0.0.1:9222` 上的浏览器
2. 读取右侧课程目录
3. 找到所有单元的 `Further listening`
4. 依次进入 `Conversation / Passage / Lectures`
5. 如果 `Lectures` 内有 `Lecture 1 / Lecture 2`，会继续逐个切换
6. 提取页面上的题目和选项
7. 保存为 Word 文档
8. 最后生成总汇总文档

运行过程中请不要关闭浏览器，也不要频繁手动点击页面。

---

## 六、输出文件说明

所有输出文件都在：

```text
output/
```

### 1. 单独分类文件

文件命名格式：

```text
Unit编号——分类——题目类型.docx
```

示例：

```text
Unit1——Further listening——Conversation.docx
Unit1——Further listening——Passage.docx
Unit1——Further listening——Lectures——Lecture 1.docx
Unit1——Further listening——Lectures——Lecture 2.docx
Unit2——Further listening——Conversation.docx
```

### 2. 全部题目汇总文件

批量爬取完成后会生成：

```text
output/全部题目汇总.docx
```

汇总文档会按以下结构分类：

```text
Unit1
  Further listening
    Conversation
    Passage
    Lectures
      Lecture 1
      Lecture 2
Unit2
  Further listening
    Conversation
    Passage
    Lectures
      Lecture 1
      Lecture 2
```

文档中只包含题目内容、Directions 和选项，不包含 JSON 元数据。

---

## 七、只爬取当前页面

如果不想批量爬取，只想爬当前打开的页面，可以执行：

```powershell
$env:CRAWL_MODE="current"
npm run crawl
```

恢复批量模式：

```powershell
Remove-Item Env:CRAWL_MODE
npm run crawl
```

---

## 八、常见问题

### 1. 报错：`connect ECONNREFUSED 127.0.0.1:9222`

原因：浏览器没有用远程调试模式启动。

解决：重新执行 Chrome 或 Edge 的启动命令：

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="<你的临时浏览器配置目录>"
```

或：

```powershell
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="<你的临时浏览器配置目录>"
```

### 2. 报错：找不到 `chrome.exe` 或 `msedge.exe`

原因：浏览器安装路径和示例不同。

解决：在电脑中搜索：

```text
chrome.exe
msedge.exe
```

找到实际路径后，把命令里的路径换成你自己的路径。

### 3. 页面跳出“我知道了”弹窗导致无法切换

脚本已经内置自动关闭逻辑，会自动点击：

- `我知道了`
- `知道了`
- `确定`
- `确认`
- `OK`

如果仍然卡住，可以手动点掉弹窗后重新运行。

### 4. 生成的题目没有听力问题文本

部分听力题的问题只在音频中播放，网页 DOM 中只显示选项。

脚本只能爬取网页上实际存在的文字，无法直接从音频中识别题目。

### 5. 旧的 `.json` 文件还在 output 里

这些是早期版本生成的旧文件，不影响现在运行。

如果不需要，可以手动删除：

```powershell
Remove-Item .\output\*.json
```

### 6. 想重新爬取一遍

直接重新执行：

```powershell
npm run crawl
```

同名 `.docx` 文件会被覆盖为最新内容。

---

## 九、推荐完整操作流程

每次从零开始，推荐按这个顺序操作。请先把 `<你的项目路径>` 和 `<你的临时浏览器配置目录>` 替换为你自己的实际路径：

```powershell
cd "<你的项目路径>\English4"
npm install
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="<你的临时浏览器配置目录>"
```

然后在新浏览器窗口登录课程网站并进入课程页面。

最后执行：

```powershell
npm run crawl
```

完成后打开：

```text
output/全部题目汇总.docx
```


