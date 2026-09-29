
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
Private Const CFG_ASM_SELF      As Boolean = False  ' 展开时，装配体本身（以及各级子装配体）是否也各出一份
                                                    ' 默认 False：只交付零件 —— 装配体/子装配体自身
                                                    ' 以及它们的工程图都不导出（想要就改 True）

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

'---- 1.7 装配体展开时的"参与状态"过滤（隐藏 / 压缩 / 封套，只作用于"装配体展开"）----
'  这三种组件在装配体里都不参与当前的设计状态，导出它们只会污染交付物：
'    隐藏 —— 视觉上根本看不到（Component2::IsHidden）
'    压缩 —— 不参与重建、也不进 BOM（Component2::IsSuppressed）
'    封套 —— 只是空间占位的参考体，不是真实零件（Component2::IsEnvelope）
'
'  判定会沿"父级链"向上走：某个子装配体被隐藏/压缩时，它下面的零件同样跳过
'  （子零件自身往往是"显示"的，但实际跟着父级一起看不见）。
'
'  ！可见性随【当前激活的配置】变化 —— 这里过滤的是运行那一刻的实际状态。
'     同一个装配体换个配置再跑，结果可能不同，这是符合预期的。
'  ！轻化（Lightweight）不算：那只是加载方式，零件本身有效，照常导出。
Private Const CFG_SKIP_HIDDEN     As Boolean = True ' 跳过 已隐藏 的零件
Private Const CFG_SKIP_SUPPRESSED As Boolean = True ' 跳过 已压缩 的零件
Private Const CFG_SKIP_ENVELOPE   As Boolean = True ' 跳过 封套 零件
'  被跳过的组件不会静默消失：清单里记"跳过"并写明是哪种状态（含"上级装配体"字样）

Private Const CFG_ASK_SKIP_OPTIONS As Boolean = True
'  True ：装配体展开后弹一个勾选框，把上面三类**逐行列出**（带各自的实际个数），
'          由用户当场决定要哪几类。勾选框的**初始状态取自上面三个 CFG_SKIP_***
'          —— 比如 CFG_SKIP_HIDDEN = False 时，"也导出已隐藏"一进框就是勾上的。
'          用户在框里改的只影响这一次运行，不写回配置。
'  False：不询问，直接按上面三个 CFG_SKIP_* 的设定执行
'  只有装配体才会弹（零件 / 工程图没有"隐藏的零件"这回事）。
'  万一勾选框弹不出来（PowerShell 被安全策略拦了之类），会按上面的配置继续，不会卡住。

'---- 1.8 其它 ----
Private Const CFG_WRITE_REPORT  As Boolean = True   ' 是否在输出目录生成导出清单（CSV，Excel 可直接双击打开）

Private Const CFG_REPORT_WITH_SRC As Boolean = True ' 清单文件名带上"来源名"，便于一眼看出是哪个文档导出的
Private Const CFG_REPORT_SUFFIX   As String = "_导出清单"
'  清单文件名 = <来源名><CFG_REPORT_SUFFIX>.csv
'    来源名：当前文档模式 = 文档名（如 家庭服务机器人-重组）
'            文件夹模式   = 源文件夹名
'            多选文件模式 = "多选文件"
'  例：家庭服务机器人-重组_导出清单.csv
'  CFG_REPORT_WITH_SRC 改成 False 就回到原来的固定名 _导出清单.csv
Private Const CFG_USE_PS_DIALOG As Boolean = True   ' 多选文件用 PowerShell 原生对话框
Private Const CFG_QUIET_OPEN    As Boolean = True   ' 静默打开文档（不弹配置 / 参考丢失对话框）

Private Const CFG_MODERN_FOLDER_PICK As Boolean = True
'  选文件夹时用哪种对话框：
'    True ：Windows 通用对话框（和 SolidWorks「另存为」同款）——
'           左侧有"快速访问"，顶部地址栏能直接**粘贴路径**，可新建文件夹
'    False：旧式树形对话框（BrowseForFolder）
'  True 时如果脚本被安全策略拦住，会自动退回旧式对话框，不会卡住

Private Const CFG_VER           As String = "v2.3.0" ' 版本号（只用于弹窗标题显示）

'---- 1.9 上次目录记忆（存于 HKCU\Software\VB and VBA Program Settings\SWauto\Export）----
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
