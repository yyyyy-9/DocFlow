# 模板样式工具

把一份DOCX中的报告样式设置到带封面、目录和页眉的DOCX模板上，输出新模板。GUI和CLI共用原template_prepare，不启动Word/WPS，不覆盖原件。

安装Python依赖：`python -m pip install -r requirements.txt`，然后运行 `python template_tool.py`。Windows图形界面需要Python的Tk支持；EXE已自带运行环境。读取模式可选择内置或已能读同一文件的Python。

先选择样式来源（默认内置母模板0.3.11），选择带封面的目标模板，读取内容并明确正文起点，选择新输出路径，再生成。正文之前的封面/目录/页眉/节保留；正文起点通过普通空段落中的WPS_ReportBody书签定位。生成后的模板可在HRS的Word导出窗口选择，目录随后由Word/WPS更新。

CLI只读检查：`python template_tool.py --inspect 路径.docx`。写入先用检查结果取得源SHA，`--prepare`须明确`--body-start`、`--expected-sha256`和新`--output`；`--style-source`可指定另一份样式来源。使用`--help`查看完整参数。

测试：`python -m unittest discover -s tests -p "test_template_*.py" -v`。输入均由测试现场生成，不使用商业报告。Windows构建：安装requirements-build.txt，运行`python tools/build_template_tool.py --out 新输出目录 --work 新构建目录`。

当前工具版本1.0.3；公开包复用既有EXE字节，附同源源码/许可证。src/wps_pipeline同时由HRS可选Word组件构建器复用，自有代码按仓库GPL-3.0-or-later授权，lxml、Python、Tcl/Tk、PyInstaller等保持原许可。

公开附带的cover_template_v0311.docx为项目自造的通用封面/目录/页眉示例，可由tools/build_example_cover.py从格式母模板重新生成；没有使用外来封面或个人元数据。格式母模板和转换代码保持原字节。
