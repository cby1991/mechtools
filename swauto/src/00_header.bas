Option Explicit

'==========================================================================================
'  SWautoExport
'  SolidWorks 一键导出：模型 -> STEP，工程图 -> PDF
'------------------------------------------------------------------------------------------
'  ！版本号不写在这里 —— 唯一来源是 10_config.bas 的 CFG_VER（弹窗标题、导出清单、
'    SwAuto_Help 都直接引用它）。这个文件头以前写死过 "v2.2.0"，结果一路漂了 5 个版本
'    没人发现（v2.4.2 时它还写着 v2.2.0）。不要再往这里加版本号。
'    自检有 `test_cfg_ver_is_single_source` 盯着"全项目只有一处 CFG_VER 定义"。
'------------------------------------------------------------------------------------------
'  ★ 平时只用这一个：SwAuto_ExportActiveDoc  （模式 1/3）
'      （打开要导出的零件/装配体 -> 运行它 -> 选输出目录 -> 完成）
'
'  运行入口（工具 > 宏 > 运行，建议绑定快捷键或工具条按钮）：
'      SwAuto_ExportActiveDoc   【模式 1/3 · 当前文档】【主入口】
'                                 零件   -> 本零件 STEP + 对应工程图 PDF
'                                 装配体 -> 只导它下面的零件 STEP + 各自工程图 PDF
'                                           （装配体/子装配体本身不出，配置 1.5 可改）
'                                           （螺钉螺母等标准件自动过滤，配置 1.6）
'                                           （隐藏/压缩/封套：弹勾选框逐类询问，配置 1.7）
'                                 工程图 -> 本图 PDF + 它引用的模型 STEP
'      SwAuto_PickFiles         【模式 2/3 · 自选文件】多选文件，想只导几个时用
'      SwAuto_PickFolder        【模式 3/3 · 文件夹批量】整个文件夹批量导出
'      SwAuto_Help              【使用说明】宏一览 + 快捷键绑定方法 + 当前打开的文档
'      SwAuto_CheckTag          【自测】检查当前文档有没有带"跳过标记"（打完标记想确认时跑它）
'------------------------------------------------------------------------------------------
'  ！绑快捷键时的大坑（已经踩过一次）：
'     工具 > 自定义 > 键盘 > 类别选「宏」>「新建宏按钮」里有一个【方法】下拉框，
'     因为本模块含多个入口，**必须**把「方法」选成 SwAuto_ExportActiveDoc。
'     漏掉这一步，SolidWorks 会挑一个方法运行 → 按快捷键却弹出「文件夹批量模式」。
'     每个模式的所有弹窗标题都会写明自己是几号模式，运行起来一眼就能看出有没有绑错。
'------------------------------------------------------------------------------------------
'  所有可调参数集中在【第 1 节 配置区】，改完直接保存即可，无需重新导入。
'  每次运行都会在输出目录生成 _导出清单.csv，并弹窗汇总（含"未找到工程图"和被过滤的标准件）。
'==========================================================================================

'---- 全局状态（单次运行内有效）----
Private m_swApp          As SldWorks.SldWorks
Private m_OutRoot        As String           ' 用户选择的输出根目录
Private m_StepDir        As String           ' STEP 实际输出目录
Private m_PdfDir         As String           ' PDF  实际输出目录
Private m_UsedNames      As Object           ' Dictionary：本次运行已占用的输出路径(大写)
Private m_DoneSrc        As Object           ' Dictionary：源文件规范化路径 -> 已生成的输出文件（同一文件只导一次）
Private m_Results        As Collection       ' 结果行：Array(类别, 源文件, 输出文件, 结果, 备注)
Private m_NoDrawing      As Collection       ' 未找到工程图的模型路径
Private m_SkippedStd     As Collection       ' 被当作标准件过滤掉的零件路径
Private m_SkippedState   As Collection       ' 因"隐藏/压缩/封套"被跳过的组件标识（同一次运行内去重）
Private m_SkipWhy        As Object           ' Dictionary：被跳过的组件标识 -> 跳过原因（与上面同步维护）
Private m_TagCache       As Object           ' Dictionary：文件路径(大写) -> 是否带「跳过标记」（同一文件只查一次属性）

'---- 复用的 COM 对象（懒加载，见 20_util.bas 的 U_Fso）----
' Scripting.FileSystemObject 的创建走 COM 激活，而"文件是否存在"在装配体展开时
' 会被问成千上万次，所以建好一次就一直用，不要每次调用都 CreateObject。
Private m_fso            As Object

'---- 装配体展开时，"不参与当前状态"的组件按类别暂存（等用户在勾选框里拍板）----
' 每个元素是 Array(显示标识, 磁盘路径, 原因文字)。磁盘路径为空 = 拿不到路径，勾了也导不出来。
' 之所以要"先存后定"：勾选框里要显示每一类的实际个数，而这个数只有扫完才知道；
' 用户勾选后又要把对应的那类放回待导出列表 —— 所以不能一边扫一边就把它写进"跳过"。
Private m_ListHidden     As Collection       ' 已隐藏
Private m_ListSuppressed As Collection       ' 已压缩
Private m_ListEnvelope   As Collection       ' 封套
Private m_SrcName        As String           ' 本次导出的"来源名"，用于给导出清单命名（文档名 / 文件夹名 / "多选文件"）
Private m_OriginalActive As String           ' 运行前的活动文档标题，跑完还原
Private m_ModeTag        As String           ' 当前模式的标识，如 "1/3 · 当前文档"；写进所有弹窗标题，防止绑错宏后看不出来

'---- 工程图索引（惰性构建，批量时只建一次）----
Private m_IdxByPath As Object                ' Dictionary：模型全路径(大写) -> 工程图全路径
Private m_IdxByName As Object                ' Dictionary：模型文件名(大写) -> 工程图全路径
Private m_IdxBuilt  As Boolean
