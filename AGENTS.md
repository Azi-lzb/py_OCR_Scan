# ScanToExcel — Agent 工作须知

本地离线 OCR 工具：拍照/扫描的会计月计表 → Excel（勾稽校验）→ 宽表汇总 → 录入系统。
架构 = core/（唯一业务核心）+ shell-pywebview/（桌面壳，主力）+ shell-flask/（浏览器壳，备用/未来 Linux）。

## 铁律

- **打包和 git 只在用户明确要求时做**（两者都耗时，用户多次强调"不要随便打包/git"）。
- 数据唯一目录 = `core/data/`（config.xlsx、templates/Template.xlsx、transforms/国库报表.xlsx、exports/）。
  打包种子"不存在才复制"，绝不覆盖用户修改；根目录不该有 data/。
- 用户业务文件在 `external/`（真实照片、手工表），已 gitignore，**绝不能入库**。
- 改完后端必须重启 8750 服务（用户页面常连着旧进程，"没反应"多半是这个）。
- 用户当前主线流程：OCR 识别 → 导出 Excel（按标题）→ 数据处理 → 宽表汇总+写入目标。
  用户需对旧批次照片**重新识别+重新导出**才能拿到新 sheet 命名与几何修复效果。

## 打包（双壳 onedir，当前 380.0 / 377.9 MB）

- `shell-pywebview/build_exe.py`、`shell-flask/build_exe.py`：PyInstaller onedir，`--noconsole`（flask 保留控制台）。
- 必带：`--collect-all webview/rapidocr_onnxruntime/rapid_table/pypdfium2/pillow_heif`、
  `--hidden-import webview.platforms.edgechromium`、**`--hidden-import tkinter`（两个壳都要，
  剔了它 Tk 回退路径会报 module not found）**。
- 排除：COM/pywin32 全家（无 Excel 自动化）、pywebview 版另排 qt/gtk/cocoa/android。
- 种子装配：config.xlsx、templates/Template.xlsx、transforms/国库报表.xlsx——copy-if-missing。
- 精简话题已关闭（用户拍板"不用精简，没有浪费的就行"），server 模型/pillow_heif 都在用，别再提。
- dist 与仓库对应：仓库推到 main 的提交 = 当时打包的代码；改码后 dist 即旧，等发版再重建。

## 文件对话框（2026-10-10 根治，别走回头路）

桌面壳（pywebview）**不要用 pywebview 的 create_file_dialog 路径**，用 core 里的
`_sta_open_dialog / _sta_save_dialog`：专用 **STA 线程**（pythonnet Thread + SetApartmentState）
直接跑 WinForms OpenFileDialog/SaveFileDialog。原因（探针实测）：
1. pywebview `parse_file_type` 的正则不接受描述里的 `/`（"图片/PDF"直接 ValueError）——
   **过滤器描述只能用 \w 和空格**（现用"图片或PDF"）；
2. 冻结态 js_api 线程是 MTA，ShowDialog 直调/Form.Invoke 均挂起。
桌面壳原生失败**不回退 Tk**（后台线程 Tk 会崩整个进程），抛可见 RuntimeError。
Flask 壳（无窗口）走 `_tk_*` 专用线程执行器，一直正常。
设置页有四模式：自动/系统原生/Tk 对话框/浏览器内置（浏览器内置 = 前端 input 选文件 →
base64 → `upload_files` 落盘 data/uploads → 显式路径进 choose_images/dp_pick）。
浏览器模式下保存类导出不弹窗，落到 data/exports 默认名。

## 模板体系要点

- Template.xlsx 页名：`一页`/`二页`（隐藏表 `几何` 勿动）。改名必须同步三处：
  命名区域 attr_text、几何 JSON 键（已有孤儿键自愈逻辑）、确认无代码写死页名。
- 几何比例 = 程序在建模板/首次完整命中时采样的"完美切割图"（行列边界百分比），
  属于模板页不属于照片；OCR 新图直接复用，不为每张生成。
- 导出 sheet 命名（title 模式）：`表格标题#模板页名`，同标题同页组加 `#1/#2`，
  超 31 字符先压标题里的"（国库）"；名字必须含"会计月计表"（宽表汇总关键字靠它命中）。

## 宽表汇总 / 目标写入（pytools 3.9.8 同口径，勿动语义）

- 复刻类需求的铁律：**读参考源码移植 + 对参考输出逐行穷举比对**，不能按列名语义自己写
  （教训：8 处口径偏差 + 198/204 行差异拖到用户人工发现）。
- 目标写入 `append_to_target`：去重键日期统一取日期部分（date vs datetime 文本化不对称会重复追加）；
  表头缺前缀列保护性跳过；目标被 Excel/WPS 占用报 PermissionError 要转可读提示。
- 用户目标簿 = `core/data/transforms/国库报表.xlsx::源数据时序`（用户自管，纯数据簿）。

## Linux 迁移清单（用户计划中，动手时照此做）

- 纯 Python 核心（OCR/模板/导出/宽表）零改动，依赖全有 Linux 版。
- 需适配三处：① `open_path` 的 `os.startfile` → `sys.platform` 分支调 `xdg-open`；
  ② `single_instance.py` Windows 互斥体 → Linux 跳过；③ 对话框 → pywebview GTK 原生
  或「浏览器内置」模式（STA WinForms 路径加平台守卫）。
- 环境：Tk 需 `python3-tk`；pywebview Linux = WebKitGTK（系统包 libwebkit2gtk）；
  PyInstaller 不能交叉打包，须在 Linux 上打 Linux 包。

## 调试方法论（本项目反复踩过的坑）

- 后端 API 测试：Flask `create_app().test_client()` 免端口；POST body 是**裸 JSON 数组**
  （`handler(*args)` 展开），单个数组参数才需双层 [[...]]。
- 前端改动必跑 `tests/check_frontend_js.py`（node --check）；"跨行字符串 WARN"是既有误报可忽略。
- 打包版冒烟：pywebview exe 无外部接口，需 SetProcessDPIAware 统一物理坐标 + 截图定位 +
  mouse_event 真实点击；exe 的数据在 **dist/ScanToExcel/core/data**（不是项目根）。
  **注意 Flask 服务的 Tk 对话框会弹在屏幕上拦截点击**，测试前先 ctypes EnumWindows 清残留。
- 用户正被 WPS/Excel 编辑的文件（~$ 锁文件）不要动。
