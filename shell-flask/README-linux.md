# Linux 发行包

支持 x86_64、glibc 2.28 及以上的 Linux。使用本机浏览器界面，内置 Python、OCR 模型、PDF 与 Excel 依赖，无需安装 Python 或联网下载模型。

[下载 Linux 1.0.20261113](https://github.com/Azi-lzb/py_OCR_Scan/releases/tag/v1.0.20261113-linux)。公开发行包清空了模板金额和目标簿历史明细，保留规则、模板几何与目标表头。

## 便携版

```sh
tar -xzf ocr-guoku-accounting-1.0.20261113-linux-x86_64-glibc2.28.tar.gz
cd ocr-guoku-accounting-1.0.20261113-linux-x86_64-glibc2.28
./start.sh
```

启动器通过 `xdg-open` 打开默认浏览器。没有桌面环境时使用 `./start.sh --no-browser`，按终端显示的地址访问；服务仅监听本机，默认端口 8750，被占用时自动选择 8751–8769。

## Debian / Ubuntu 安装版

```sh
sudo apt install ./ocr-guoku-accounting_1.0.20261113_amd64.deb
ocr-guoku-accounting
```

也可从应用菜单启动「OCR 国库会计工具」。系统需提供浏览器和 `xdg-utils`；安装包声明 `libc6 >= 2.28` 和 `xdg-utils` 依赖。

## 数据与功能

两种包都将用户数据保存在 `${XDG_DATA_HOME:-$HOME/.local/share}/ocr-guoku/core/data/`。配置、模板、目标工作簿只在不存在时从包内复制，升级或重启不会覆盖已有文件。导出文件位于该目录的 `exports/`。

种子规则中指向 Windows `core/data/transforms/国库报表.xlsx` 的路径在运行时映射到上述用户数据目录；配置文件本身保持原样。自定义 Windows 绝对路径需要在规则配置中改为 Linux 路径。

Linux 使用浏览器文件选择。OCR、PDF 输入、Excel 导出、勾稽校验和数据归集使用共同核心；依赖 Windows COM 的 Excel/WPS/Word 格式转换在 Linux 不可用。Excel 文件可以交给系统已安装的兼容办公软件打开。

## 从源码构建

在仓库根目录运行（需要 Linux 容器模式的 Docker）：

```sh
docker build -f shell-flask/Dockerfile.deb --build-arg APP_VERSION=1.0.20261113 --target artifact --output type=local,dest=shell-flask/dist/linux-glibc2.28 .
```

构建在 manylinux glibc 2.28 环境进行，补充匹配的共享 libpython，收集应用需要的非 glibc 动态库，并逐个检查 ELF 的 GLIBC 版本要求。产物包含 `.deb`、`.tar.gz` 和 SHA-256 校验文件。Windows 发行目录不会被覆盖。

本地打包种子可能包含用户数据。公开分发前，在独立目录准备公开包（保留原包）：

```sh
python shell-flask/prepare_public_release.py --source shell-flask/dist/linux-glibc2.28 --output shell-flask/dist/linux-public-1.0.20261113
```

该步骤只替换三个种子工作簿：模板数字区域清空金额、目标簿只保留「源数据时序」表头、配置中的默认目标路径改为通用路径；程序二进制与模型保持原样。发布输出目录中的安装包与对应 `SHA256SUMS`，不要上传本地原包或 `core/data/`。
