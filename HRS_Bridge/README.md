# HRS ↔ WPS Bridge 0.4.3

独立的本地编排与转换工具。HRS 2.0 提供 Markdown/对象/页面模型；原三表 0.1.6 提供 Word 扫描、识别、局部复核、模板装配和独立检查。Bridge 调用这些现有入口，不另建识别器或模型 API 客户端。

## 本次版本与依赖

- Bridge 唯一身份源为 `hrs_bridge/__init__.py`；`package.json`与根目录COMPONENTS.json由`tools/maintain_manifest.py`同步。
- DocFlow源码固定引用HRS2.0.7；整体下载包与独立Bridge0.4.3包使用同一HRS App和页面模型。实际源码引用与附件组成见release.json、COMPONENTS.json及包内清单，独立包输入另见SOURCES.json。document-v2、公式编号/引用和数学适配保留。
- 三表0.1.6 / 内核协议0.5.7只保留一份共享源码。公开构建身份见BUILD_INFO与COMPONENTS；本次内核构建仅增加原生链接范围事实与独立核对，原编号事实校验继续执行，正文识别/编号规则保持，公开文档和通用封面模板随发行提供。`forms_tool`和`report_tool`共用这一份内核。
- UI0.1.2与Bridge共享内核，各自控制任务分支。HRS子模块固定到本次公开2.0.7提交；已有冻结任务继续使用匹配的程序与指纹，不自动迁移。

Python 3.12+。Clean 包包含实现和必要说明；DocFlow Full增加当前测试/自造输入、CPython3.12 Win64离线wheel与playwright-core。Python、Node、Chrome、Pandoc、WPS 不重复嵌入。

在发行根目录：

```powershell
python verify.py
python -m pip install -r requirements-browser.txt
# Full 离线安装：
python -m pip install --no-index --find-links offline/wheels -r requirements-browser.txt
python run.py bridge --version
```

Word→Word 三表只需 Python 及其文档依赖；使用原生 WPS 页面读取或旧 DOC 转换时，另需本机 WPS/pywin32。HRS 浏览器 CLI 另需 Node20+、Chrome 和 playwright-core；可用当前 HRS 页面模型完成 Word 输出而无需再起浏览器。原生公式转换仍由单独配置的 Pandoc 完成。

源码运行时，在本目录用 `python -B -m hrs_bridge ...`。以下命令使用DocFlow根目录的`python run.py bridge ...`；独立Bridge组件包使用自己的bridge.py入口。全局根参数写在子命令前：`--hrs-root`、`--wps-root`、`--forms-root`，或设置 HRS_ROOT、WPS_ROOT、WPS_FORMS_ROOT。浏览器/Node/Pandoc 分别用 CHROMIUM_PATH、HRS_BRIDGE_NODE、PANDOC_PATH；外置 Node 模块用 HRS_BRIDGE_NODE_MODULES。

DocFlow共享HTML_Report_Studio与WPS_Forms_Workflow，读取根目录 `integration/bridge_runtime.json`。显式路径优先，不依赖历史版本目录。不能在同一 Python 进程混载两份 HRS/WPS 模块。

## 原生Word超链接导入

原生Word超链接（w:hyperlink）保留显示文字、加粗等强调及可确认目标。HTTP/HTTPS、邮件和绝对file地址沿原链接入口处理；指向已确认章节、图表或其题注的唯一书签映射到HRS稳定身份，预览、HTML和再次导出Word继续可定位。未知/重复书签、危险或相对地址保留文字与原目标记录，标为NEEDS_REVIEW，不激活不可靠链接；不猜新的路径基准。HYPERLINK等复杂域、嵌套链接仍按原受限导入边界处理。

HRS入口及核对窗口标记beta；from-docx命令协议不变。只变更链接目标也会进入原生语义指纹，避免误恢复旧Markdown。新增纯数据链接专项纳入原core/bridge入口，不依赖浏览器；HRS实际页面、HTML与Word往返另列。

## 新推荐路线：三表连续阶段稿

新建 job.json，路径相对于清单目录，输出必须是新目录：

```json
{
  "schema": "hrs-bridge-job-v2",
  "output": "output/run01",
  "tasks": [
    {
      "id": "reports",
      "source": {"kind": "wps-forms", "path": "input"},
      "outputs": ["docx"],
      "title": "试验报告汇总",
      "template": "templates/prepared.docx",
      "forms_options": {
        "policy": "Agent辅助核查",
        "pages": "off",
        "doc_conversion": "off"
      }
    }
  ]
}
```

```powershell
python run.py bridge run-job job.json
python run.py bridge job-status output/run01
python run.py bridge forms-review-plan output/run01 --task reports --fraction 0.2
```

运行顺序是：生成01输入表 → 原程序扫描并立即交首稿 → 填当前02分类 → 提交分类新稿 → 填当前03的一批局部反馈 → 提交完整累积新稿。`pages`、`doc_conversion` 可选 off/wps，默认分别为 off/wps；没有原生WPS环境时用上例的 off。policy 只接受“人工核查”或“Agent辅助核查”；不在首稿前自动接受规则建议。模板可省略以使用原配置，任务模板优先于清单模板。

首轮的当前表单位于：

```text
output/run01/tasks/reports/rounds/001-first/02_扫描核查.xlsx
output/run01/tasks/reports/rounds/001-first/03_反馈提交.xlsx
```

填写原表要求的确认项（02 B6“确认本表”，03 B6“提交本表”），然后提交：

```powershell
python run.py bridge forms-submit output/run01 --task reports --operation classify --form "output/run01/tasks/reports/rounds/001-first/02_扫描核查.xlsx" --origin user
# 分类成功后，从返回的 forms_current 取得新稿的03；不可继续提交旧03。
python run.py bridge forms-submit output/run01 --task reports --operation feedback --form "output/run01/tasks/reports/rounds/002-classify/03_反馈提交.xlsx" --origin model
```

新轮次的全部DOCX哈希、状态/清单/待审信息和表单声明先在临时结构内核对，全部成功后才一起切换当前轮次。临时读取失败或新稿哈希不符时，当前稿、可编辑表、交付及历史保持上一轮；只登记拒绝的attempt和新目录现场。解除故障后可再次显式提交，使用新的尝试目录；程序不自动重跑整链。

每次提交复制并绑定 XLSX 快照，再调用原 classify/feedback 校验一次；不修改冻结 spec、不跳过原表确认，也不直接操作托管的内核子任务。成功后返回新的 forms_current、当前02/03、各完整 DOCX、pending_count、source_gaps、forms_rounds 和独立原回执。失败尝试记录保留，当前已交稿仍可用；重提请按实际回执取得下一路径。

只允许修改**当前一轮02/03的字节**；旧表、冻结输入、模板、历史/当前DOCX与其他任务产物受完整性检查。原件、运行代码或源快照变化时拒绝接续，必须创建新任务。完整阶段稿与所有旧稿保留，没有默认“两稿上限”。

`forms-review-plan` 只读地按原文章顺序聚合待审项。fraction 约束每批待审项数量的目标比例，整篇文章不切开，大文章会超过目标；派生拆分单元回到原来源。每批提交后对新稿重新取计划，不能拿旧快照的数量冒充当前进度。程序导出的未决项与用户尚未提交的新添反馈分开，后者不计入计划。Agent 只读必要局部证据并填写判断，程序自身模型调用数为0。

UI 0.1.2 与 Bridge 共用原表单协议，但各自维护自己的批次时间线。Bridge 管理的任务经 forms-submit 接续；不自动接管 UI 在包外生成的新分支，不能让两个控制器同时接续同一批次。

## 既有多格式与原生入口

| 输入 kind | 路线 |
| --- | --- |
| wps-forms | 来源根目录 → 原三表首稿/分类/反馈累积稿，只输出 docx |
| docx + docx | 原 report_tool 原生Word→Word，不经过Markdown |
| wps-job | 原WPS多来源清单；make_master=true才生成总稿 |
| docx + hrs/html | 原受限Word→HRS识别/投影，需要时局部结构确认 |
| markdown / hrs | HRS实际源文/资源冻结，原解析与输出策略 |
| html | 原HRS可回编归档恢复；纯阅读HTML不伪装源文 |

旧 docx/wps-job 默认 review_policy=pause。日常首稿应显式 retain-unresolved：交保留未定内容的草稿和问题清单；PARTIAL/COMPLETE 原地不会接受 job-review 或被 resume 重建。旧暂停任务用 job-review + resume-job；需要连续分类/反馈新稿的新任务使用 wps-forms。不要把三个入口的生命周期混用。旧 v1 清单兼容入口保持原行为。

任务按声明顺序执行，独立任务失败隔离。一个 wps-job/wps-forms 内的原规则可以合稿；多个独立 docx 任务不会自动合并，Markdown/Word混合总稿尚未实现。WPS原生合稿来源名占H1、源H5进入H6；原生Word图表全文SEQ编号、标题与列表编号实例隔离、明确书签REF保留，**该原生链公式暂不编号**。

HRS→Word 则遵守 HRS2 文档选项：工程章节/图表编号，公式两模式均按 show_equation_numbers；普通Markdown表格与增强表格共用表序列，Mermaid共用图序列。公式显示号启用时使用原生 SEQ公式/REF，关闭时保留公式身份和可定位引用。两条路线的不同公式策略不能互套。

```powershell
python run.py bridge recover-html input/report.html --out restored
python run.py bridge export-html restored --out output/report.html
python run.py bridge read restored --action objects
python run.py bridge to-docx restored --template templates/prepared.docx --out output/report.docx
python run.py bridge from-docx output/report.docx --out back
python run.py bridge analyze-word input/external.docx
```

HRS工程入场消费原A0 sourceContractValid，兼容旧同范围字段；页面Marked、脚注、资源和Word装配继续检查，不把结构检查当完整预检。多个H1、章节跳级阻止工程输出；手工章号保留为文本并提醒，显式 --accept-numbered-headings 只接受当前源快照。派生Word来源元数据不能绕过实际工程检查。

## 模板、编号与受限回编

模板必须先经原模板工具准备正文槽位。使用普通空段落和 WPS_ReportBody 书签，不造正文大框；复用原 inspect/open 接口保留封面、目录域、页眉页脚和节设置。映射只覆盖生成正文。

单篇HRS H1是无编号Title，H2—H6对应Word标题1—5。外部Word第六级按既有规则转列表，有效的源摘要/UTF-16范围提示可恢复其原级别。标题与正文列表使用独立编号实例，不靠改写数字伪装自动号。图表题使用 SEQ图/表 和数字书签REF；题名迁出完整紧邻旧号，技术数值保留。手改Word后按F9更新域；原生WPS实际编辑与XML检查分开验收。

组图每逻辑排为独立图片行/子图题行，同排题注起点对齐，保持宽高比、顺序、身份和外题注，兼容旧同单元格映射。外来Word表内有图且无可靠表题按图/组图，有表题按数据表；使用原WPS figure_layout，不再单写规则。受HRS最多3行/4列/12图及单行说明的表示范围约束，合并/稀疏/单格多图等无法安全表示时拒绝并保留源。

输出可选 --notes-mode footnotes/references，参考文献章节进目录不占章号；--strip-body-emphasis 只清正文额外强调，数学/模板格式保持。脚注、NOTEREF、题注安全内联与原生OMML复用现行实现。复杂未知HTML或域、未知数学不降成图片或源码冒充成功。旧式分组 rm 仅在Pandoc转换副本中等价归一，源文不改。

生成DOCX伴随 `.docx.hrs-bridge.json`，只供回HRS恢复源身份/未变片段；Word独立打开编辑不依赖此文件。另存新名后用 --provenance 显式选原映射。修改内容从当前DOCX读取；新增不能表示的结构、丢失映射、异常身份/域明确拒绝。未知Word任意排版的无损恢复不在承诺范围。

外部Word空 proofErr/lastRenderedPageBreak 只忽略原XML中确认无正文的标记；新扫描器 layout_marker 与旧 unsupported 表示均处理，其他未知内容仍拒绝。Title缺失而首个正文来自图片说明等极小样稿，来源题名识别可能需要局部判断，不以版式测试样本代替通用准确率。

## 受控块 plan/apply

```powershell
python run.py bridge block-plan restored --id tbl-results --expect-sha256 <源SHA> --text-file replacement.txt --out candidate.md --plan-out edit.json
python run.py bridge block-apply edit.json --expect-plan-sha256 <已审计划SHA>
```

plan只读，复用HRS原定位和校验；operation支持replace/insert-before/insert-after。apply重新读取并验证源、候选、修改范围、资源，复用原锁/备份/原子保存；旧快照、越界、缺资源或占用明确失败，不用整链重试覆盖。未开放任意批量附件写入。

## 状态与验收

CLI stdout 为JSON：完成退出0，拒绝/失败1，需判断或局部保留2。WORKFLOW.json为完整回执，DELIVERY.md为当前交付索引，finish_list.json列未完成任务。各原三表状态/交付/局部保留与来源缺口保留，不因存在DOCX或退出0就称完整。COMPLETE只表示本次适用本地检查完成，WPS/Word页面和人工操作单列。

在DocFlow根目录按tests/README.md运行 `python tools/run_tests.py --group all --out .work/check01`；浏览器专项需上述真实运行依赖及本机服务能力。tests/test_forms_workflow.py验证真实原CLI四轮、累积修改、文章顺序、旧表/冻结源/运行代码/受保护输出拒绝和失败恢复；tests/test_numbering_output.py验证HRS2公式开关与普通/增强表同一序列。既有标题/列表、六级投影、模板、脚注、组图、回编与HTML测试继续复用。测试输入是固定/合成样本，不是未知报告准确率。

根测试入口统一设置共享内核路径，并生成公开HTML样本；不要直接在缺少准备步骤的目录运行全部模块。缺浏览器或依赖是未满足测试条件，不能把SKIP/失败记成通过。真实WPS页面及原生编辑验收单独进行。

## 同源构建与更新

DocFlow整体发行使用根目录`tools/build_release.py`，统一包含界面和一份共享内核。Bridge独立组件仍可使用本目录`tools/export_hrs_component.py`生成/核对快照；`tools/build_packages.py`接受明确的HRS源ZIP、三表Clean和依赖构建独立包；`tools/build_word_component.py`按CP313 wheel锁生成HRS Word组件。源码须先提交，清单记录Git身份及逐文件SHA，生成目录不独立改代码。

可选Word组件只在全新目录解压，然后在 HRS 的 Word支持设置中选择该目录；Pandoc仍单独选择。它不覆盖 HRS App/草稿/工程。已经运行中的任务继续用原程序和原指纹；新版任务使用新目录，不能改回执哈希迁移，也不自动重写损坏的旧回执。

HRS与vendor、WPS组件及第三方依赖保留各自版权/许可声明，见根目录THIRD_PARTY_NOTICES.md。Full保留wheel和playwright-core随附许可；外部办公软件不随包分发。
