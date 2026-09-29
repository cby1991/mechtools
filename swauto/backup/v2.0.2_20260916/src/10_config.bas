
'==========================================================================================
'  第 1 节  配置区  ——  按需修改
'==========================================================================================

'---- 1.1 输出格式 ----
Private Const CFG_STEP_AP       As Long = 203       ' STEP 协议：203 或 214（swStepAP 直接接受这两个值）
Private Const CFG_STEP_EXT      As String = ".STEP"
Private Const CFG_PDF_EXT       As String = ".PDF"

'---- 1.2 输出目录结构 ----
Private Const CFG_USE_SUBFOLDER As Boolean = False  ' False：STEP 与 PDF 都放进你选的同一个目录（推荐）
                                                    ' True ：在输出目录下自动建 STEP\ 与 PDF\ 两个子目录
Private Const CFG_OVERWRITE     As Boolean = True   ' True：覆盖同名文件；False：跳过并在清单里标注

'---- 1.3 工程图查找 ----
Private Const CFG_FIND_DRAWING   As Boolean = True  ' 模型导出后，是否自动找它的工程图一起导出 PDF
Private Const CFG_DRW_ALSO_MODEL As Boolean = True  ' 处理工程图时，是否把它引用的模型也导出 STEP
Private Const CFG_DEEP_SEARCH    As Boolean = True  ' 同名找不到时，扫描目录反查引用关系（慢但准）
Private Const CFG_DEEP_LIMIT     As Long = 1500     ' 深搜最多扫描多少个工程图，防超大目录卡死
Private Const CFG_EXTRA_DIRS     As String = ""     ' 额外搜索目录，多个用 ; 分隔，如 "D:\图纸库"

'---- 1.4 文件夹模式 ----
Private Const CFG_RECURSE_SUBDIR As Boolean = False ' 是否递归子目录

'---- 1.5 装配体展开（只影响 SwAuto_ExportActiveDoc 在"当前是装配体"时的行为）----
Private Const CFG_ASM_EXPAND    As Boolean = True   ' True ：装配体 → 展开导出它下面所有零件
                                                    ' False：装配体 → 只导装配体自己，不碰里面的零件
Private Const CFG_ASM_SELF      As Boolean = True   ' 展开时，装配体本身（以及各级子装配体）是否也各出一份 STEP

'---- 1.6 标准件过滤（螺钉、螺母、垫圈……自动跳过，只作用于"装配体展开"）----
'  判定规则（命中任意一条即算标准件）：
'    ① SolidWorks 组件自带的 Toolbox 标记（最准）
'    ② 文件路径里含 CFG_STD_PATH_HINTS
'    ③ 文件名里含 CFG_STD_KEYWORDS（按"词边界"匹配，不会把 Spindle 当成 Pin、Walnut 当成 Nut）
'  想放行某一类，把对应关键字从列表里删掉即可；整个功能不想用就把 CFG_EXCLUDE_STD_PARTS 改 False。
Private Const CFG_EXCLUDE_STD_PARTS As Boolean = True
Private Const CFG_STD_USE_API       As Boolean = True   ' 是否用 Toolbox 标记判定

' 文件名关键字（逗号分隔，中英文都可以加）
Private Const CFG_STD_KEYWORDS As String = _
    "螺钉,螺母,螺栓,螺柱,垫圈,垫片,挡圈,卡簧,弹性挡圈,开口销,销钉,圆柱销,圆锥销,铆钉,自攻," & _
    "Bolt,Nut,Washer,Screw,Stud,Rivet,Circlip,Snap Ring,Retaining Ring,O-Ring,Spring Pin"

' 国标 / 国际标准号（逗号分隔）——标准件文件名基本都带这些前缀
Private Const CFG_STD_STDCODES As String = "GB/T,GB-T,GBT,ISO,DIN,JIS,ANSI,BS-"

' 路径关键字（逗号分隔）——Toolbox 数据库的典型路径
Private Const CFG_STD_PATH_HINTS As String = "SOLIDWORKS Data,Toolbox,标准件,紧固件"

'---- 1.7 其它 ----
Private Const CFG_WRITE_REPORT  As Boolean = True   ' 在输出目录生成 _导出清单.csv
Private Const CFG_USE_PS_DIALOG As Boolean = True   ' 多选文件用 PowerShell 原生对话框
Private Const CFG_QUIET_OPEN    As Boolean = True   ' 静默打开文档（不弹配置 / 参考丢失对话框）
Private Const CFG_VER           As String = "v2.0.2" ' 版本号（只用于弹窗标题显示）

'---- 1.8 上次目录记忆（存于 HKCU\Software\VB and VBA Program Settings\SWauto\Export）----
Private Const CFG_APP     As String = "SWauto"
Private Const CFG_SEC     As String = "Export"
Private Const CFG_KEY_OUT As String = "LastOutputDir"
Private Const CFG_KEY_IN  As String = "LastInputDir"

Private Function Cfg_GetSetting(ByVal sKey As String) As String
    On Error Resume Next
    Cfg_GetSetting = GetSetting(CFG_APP, CFG_SEC, sKey, "")
End Function

Private Sub Cfg_PutSetting(ByVal sKey As String, ByVal sVal As String)
    On Error Resume Next
    SaveSetting CFG_APP, CFG_SEC, sKey, sVal
End Sub
