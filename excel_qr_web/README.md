# 二维码批量生成器 · 网页版

原桌面版（tkinter）的网页化重制。**选择 Excel 文件 → 设置参数 → 一键把指定列文本生成二维码并新增一列插入，另存为新文件**（原文件不受影响）。

- **完全在浏览器本地运行**：所有处理（读取、生成、写回）都在本地完成，**文件不会上传到任何服务器**
- 支持格式：`.xlsx` / `.xlsm`（保留原样式与多工作表）/ `.xls`（旧版，结果统一转存为 `.xlsx`）
- 纯静态网页（HTML + CSS + 原生 JS），无需后端，**可直接部署到 GitHub Pages**
- 打开即用：双击 `index.html` 或部署到任意静态托管均可

## 功能说明

| 设置 | 说明 | 默认 |
| --- | --- | --- |
| 工作表（Sheet） | 自动读取 Excel 中所有工作表，下拉选择要处理的表 | 第一个 |
| 首行是否表头 | 「是，跳过首行」从第 2 行读取；「否，从首行开始」 | 是 |
| 二维码内容列 | 生成二维码的文字来源列（显示“第N列（表头名）”） | 第 1 列 |
| 二维码大小 | 小 96px / 中 128px / 大 192px，只影响显示 | 中 |
| 自动调整 | 自动拉高行、加宽二维码列，保证完整显示 | 开 |

高级设置（「高级设置…」按钮内）：

| 设置 | 说明 | 默认 |
| --- | --- | --- |
| 自定义二维码边长 | 勾选后输入精确像素（32~1000），覆盖“小/中/大” | 关 |
| 纠错级别 | 低/标准/高/最高 ≈ 7%/15%/25%/30% 容错 | 标准 |
| 四周白色留白 | 紧凑 2 / 标准 4 / 宽松 6 个模块 | 标准 |
| 自定义起始行 | 表头占多行等特殊情况时使用 | 关 |

其他行为（与桌面版保持一致）：

- 二维码自动新增在**有内容的最后一列之后**（不覆盖原数据，表头自动补写“二维码”）
- 内容列为空的行自动跳过
- 整数型单元格自动去掉小数尾巴（如手机号不会出现 `13800138000.0`）
- 按二维码规范预留 ≥4 模块白色静区，并以整数倍渲染（模块边缘锐利、无模糊），微信「扫一扫」可稳定识别
- 二维码图片在单元格内**水平、垂直居中**（与桌面版一致）
- 内容长度上限：纠错级别 标准(M) 下最多约 2331 个字母/数字或 777 个汉字，超长会提示具体行号
- 顶栏「自定义二维码」：输入任意内容实时预览，可下载为 PNG 图片
- 右上角**调色板图标**按钮一键切换主题配色（微信绿 / 清爽蓝 / 活力橙 / 科技紫 / 翡翠青，见下表），选择自动记忆（localStorage）
- 内置数据预览：加载后展示内容列前若干行及二维码小图（表头与内容居中，超长内容以 … 截断），方便核对
- 日志逐条记录每条生成结果，进度实时显示

### 主题配色

右上角调色板图标按钮一键切换，选择自动记忆（localStorage），亮 / 暗色各自适配：

| 主题 | 亮色主色 | 暗色主色 |
| --- | --- | --- |
| 微信绿（默认） | `#07c160` | `#2cd97a` |
| 清爽蓝 | `#33a5e8` | `#4db9f0` |
| 活力橙 | `#f59b18` | `#ffae52` |
| 科技紫 | `#a26ff7` | `#b18efb` |
| 翡翠青 | `#17b8a6` | `#2dd4bf` |

## 项目结构

```
Excel_Qr_Web/
├── index.html          # 单页应用入口
├── css/style.css       # 样式（自适应亮 / 暗主题）
├── js/app.js           # 全部业务逻辑
├── lib/                # 本地化第三方库（可离线使用）
│   ├── exceljs.min.js      # 读写 .xlsx/.xlsm，支持单元格内嵌图片
│   ├── qrcode.min.js       # 二维码矩阵生成（支持中文）
│   └── xlsx.full.min.js    # 仅用于读取旧版 .xls（BIFF）
├── logo/favicon.png    # 站点图标（微信绿主题，可直接替换该文件改图标）
├── test/               # 开发用测试（不参与运行）
│   ├── smoke.node.js       # Node 冒烟测试：node test/smoke.node.js
│   ├── dom_test.html       # 浏览器 DOM 核心逻辑测试（headless Chrome）
│   ├── theme_test.html     # 主题切换按钮 / 面板 / 持久化测试
│   ├── preview_overflow_test.html # 预览超长内容截断、无水平滚动条测试
│   ├── style_probe.html    # 预览居中、徽标、主题图标样式断言
│   └── 测试数据_旧版.xls    # .xls 测试样例
├── 测试数据.xlsx       # 演示样例（5 个工作表）
└── .nojekyll           # 部署到 GitHub Pages 时跳过 Jekyll 构建
```

## 本地使用

直接双击打开 `index.html` 即可使用，无需安装任何东西、无需联网（第三方库均已本地化）。

## 部署到 GitHub Pages

> GitHub Pages 提供免费静态托管，很适合这种「无后端」的网页工具。

### 方式一：分支部署（推荐，最简单）

1. 新建一个 GitHub 仓库（Public 或 Private 均可），例如 `excel-qr-web`
2. 把本项目**全部文件**上传到仓库的 `main` 分支根目录（不要放进子文件夹）
3. 打开仓库 **Settings → Pages**
4. 在 **Build and deployment** 的 **Source** 下拉框选择 **Deploy from a branch**
5. Branch 选择 `main`，目录选择 `/ (root)`，点击 **Save**
6. 等待一两分钟，页面顶部会出现你的站点地址 `https://<用户名>.github.io/excel-qr-web/`

### 方式二：放在项目仓库的 `docs/` 文件夹

如果你想把工具作为某个已有项目仓库的一部分：

1. 把本项目文件放到仓库根目录的 `docs/` 文件夹内
2. Settings → Pages → Source 选择 **Deploy from a branch**
3. Branch 选择 `main`，目录选择 `/docs`，保存即可

### 方式三：GitHub Actions（可选）

也可以在仓库中新建 `.github/workflows/gh-pages.yml`，用 Actions 自动构建发布（适合想用子分支 gh-pages 的情况）：

```yaml
name: Deploy to GitHub Pages
on:
  push:
    branches: [main]
permissions:
  contents: read
  pages: write
  id-token: write
jobs:
  deploy:
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/configure-pages@v5
      - uses: actions/upload-pages-artifact@v3
        with:
          path: .
      - id: deployment
        uses: actions/deploy-pages@v4
```

（启用后还需在 Settings → Pages → Source 选择 **GitHub Actions**。）

### 部署后提示

- 站点是纯静态页面，访问 `https://<用户名>.github.io/<仓库名>/` 即可使用
- `.nojekyll` 文件已包含，避免 Jekyll 特殊处理
- 若提示 404，请确认文件确实在所选分支/目录下，并等 1~2 分钟

## 浏览器端限制说明（与桌面版差异）

网页版没有服务器，无法像桌面版那样“另存到任意文件夹”，因此结果以**浏览器下载**方式保存（文件名自动为 `原文件名_二维码.xlsx`）。

| 能力 | 桌面版（Python） | 网页版（本仓库） |
| --- | --- | --- |
| 支持格式 | .xlsx / .xlsm / .xls | .xlsx / .xlsm / .xls（.xls 转存为 .xlsx） |
| 多工作表 | ✅ | ✅ |
| 二维码插入新列 | ✅ | ✅ |
| 样式保留 | .xlsx 保留 | .xlsx/.xlsm 保留大部分；**.xls 不保留样式** |
| 宏（.xlsm） | 保留 | 转存后不保留宏 |
| 另存路径 | 可自选文件夹 | 浏览器下载（固定命名） |
| 文件体积上限 | 无 | 50MB（纯前端解析） |

## 开发与测试

```bash
# Node 冒烟测试（验证核心读取/生成/嵌图/写出流程）
node test/smoke.node.js

# 浏览器 DOM 测试（headless Chrome，需已安装 Chrome）
# 打开 test/dom_test.html 即可，或命令行：
chrome --headless=new --allow-file-access-from-files --dump-dom file:///.../test/dom_test.html

# 其余浏览器测试（同法打开 / 运行）：
# test/theme_test.html                主题切换（按钮/面板/5 套主题/持久化）
# test/preview_overflow_test.html     预览超长内容 … 截断、无水平滚动条
# test/style_probe.html               预览居中、徽标、主题图标样式
```

## 第三方库许可

- **[exceljs](https://github.com/exceljs/exceljs)** — MIT License
- **[qrcode-generator](https://github.com/kazuhikoarase/qrcode-generator)** — MIT License
- **[SheetJS CE](https://sheetjs.com)** — Apache-2.0 / SheetJS Community Edition（仅用于读取旧版 `.xls`）

> 本项目仅本地化引用以上库用于网页端数据处理，功能与桌面版 `Excel_Qr/` 对应。

## 与原桌面版

本仓库是 [Excel_Qr](../Excel_Qr)（Python + tkinter 桌面版）的网页化重制，界面与参数沿用桌面版约定。若需要离线 Windows exe，请使用桌面版。
