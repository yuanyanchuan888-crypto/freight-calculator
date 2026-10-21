# Windows 便携版构建

## 前提

- Windows x64。
- Python 3.11 或更新版本，带 Tkinter。
- 首次构建可访问 PyPI，以便安装项目和构建依赖。

## 构建

在项目目录运行：

```powershell
.\build_windows.ps1
```

脚本优先使用 `py -3.12`，其次使用 `python`。如需指定解释器：

```powershell
.\build_windows.ps1 -PythonExe "C:\Python312\python.exe"
```

脚本在项目内建立隔离的 `.venv-build`，从 `requirements.txt` 与 `requirements-build.txt` 安装依赖，并生成：

```text
dist\FreightCalculator.exe
```

这是便携式单文件程序，不是 MSI 或安装程序。可将 EXE 复制到普通 Windows 用户可访问的文件夹运行；用户机器不需要 Python。首次启动会由 PyInstaller 解包到临时目录，启动速度和杀毒软件扫描时间可能受机器影响。

## 验证

1. 确认构建脚本成功返回，文件存在且大小大于 1 MB。
2. 在构建机上双击运行 EXE，检查主界面正常显示。
3. 新增一条合成货物，确认计算结果正常。
4. 导出模板、填入一行测试数据，再导入并核对设置与结果。
5. 复制 EXE 到没有项目目录的单独临时文件夹后再次运行，确认不依赖源码目录或开发工具的 Python 路径。

EXE 尚未数字签名，也未通过 Windows Defender SmartScreen 信誉认证。公开分发前应由权利人复核构建物、进行多台普通 Windows 设备验收，并考虑代码签名与安装包方案。
