Option Explicit

'==========================================================================================
'  SWautoExport  v2.0.2
'  SolidWorks 一键导出：模型 -> STEP，工程图 -> PDF
'------------------------------------------------------------------------------------------
'  ★ 平时只用这一个：SwAuto_ExportActiveDoc  （模式 1/3）
'      （打开要导出的零件/装配体 -> 运行它 -> 选输出目录 -> 完成）
'
'  运行入口（工具 > 宏 > 运行，建议绑定快捷键或工具条按钮）：
'      SwAuto_ExportActiveDoc   【模式 1/3 · 当前文档】【主入口】
'                                 零件   -> 本零件 STEP + 对应工程图 PDF
'                                 装配体 -> 装配体下所有零件 STEP + 各自工程图 PDF
'                                           （标准件／螺钉螺母自动过滤，可在配置区关掉）
'                                 工程图 -> 本图 PDF + 它引用的模型 STEP
'      SwAuto_PickFiles         【模式 2/3 · 自选文件】多选文件，想只导几个时用
'      SwAuto_PickFolder        【模式 3/3 · 文件夹批量】整个文件夹批量导出
'      SwAuto_Help              【使用说明】宏一览 + 快捷键绑定方法 + 当前打开的文档
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
Private m_OriginalActive As String           ' 运行前的活动文档标题，跑完还原
Private m_ModeTag        As String           ' 当前模式的标识，如 "1/3 · 当前文档"；写进所有弹窗标题，防止绑错宏后看不出来

'---- 工程图索引（惰性构建，批量时只建一次）----
Private m_IdxByPath As Object                ' Dictionary：模型全路径(大写) -> 工程图全路径
Private m_IdxByName As Object                ' Dictionary：模型文件名(大写) -> 工程图全路径
Private m_IdxBuilt  As Boolean
