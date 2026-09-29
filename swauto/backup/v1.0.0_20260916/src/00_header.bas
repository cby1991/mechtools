Option Explicit

'==========================================================================================
'  SWautoExport  v1.0.0
'  SolidWorks 一键导出：模型 -> STEP，工程图 -> PDF
'------------------------------------------------------------------------------------------
'  运行入口（工具 > 宏 > 运行，建议绑定快捷键或工具条按钮）：
'      SwAuto_ExportActiveDoc   导出当前打开的文档（连同它的工程图 / 被引用模型）
'      SwAuto_PickFiles         弹文件对话框，多选文件后排入队列
'      SwAuto_PickFolder        弹文件夹对话框，批量导出整个文件夹
'------------------------------------------------------------------------------------------
'  所有可调参数集中在【第 1 节 配置区】，改完直接保存即可，无需重新导入。
'  每次运行都会在输出目录生成 _导出清单.csv，并弹窗汇总（含"未找到工程图"的模型）。
'==========================================================================================

'---- 全局状态（单次运行内有效）----
Private m_swApp          As SldWorks.SldWorks
Private m_OutRoot        As String           ' 用户选择的输出根目录
Private m_StepDir        As String           ' STEP 实际输出目录
Private m_PdfDir         As String           ' PDF  实际输出目录
Private m_UsedNames      As Object           ' Dictionary：本次运行已占用的输出路径(大写)
Private m_Results        As Collection       ' 结果行：Array(类别, 源文件, 输出文件, 结果, 备注)
Private m_NoDrawing      As Collection       ' 未找到工程图的模型路径
Private m_OriginalActive As String           ' 运行前的活动文档标题，跑完还原

'---- 工程图索引（惰性构建，批量时只建一次）----
Private m_IdxByPath As Object                ' Dictionary：模型全路径(大写) -> 工程图全路径
Private m_IdxByName As Object                ' Dictionary：模型文件名(大写) -> 工程图全路径
Private m_IdxBuilt  As Boolean
