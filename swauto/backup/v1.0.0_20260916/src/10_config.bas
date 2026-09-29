
'==========================================================================================
'  第 1 节  配置区  ——  按需修改
'==========================================================================================

'---- 1.1 输出格式 ----
Private Const CFG_STEP_AP       As Long = 203       ' STEP 协议：203 或 214（swStepAP 直接接受这两个值）
Private Const CFG_STEP_EXT      As String = ".STEP"
Private Const CFG_PDF_EXT       As String = ".PDF"

'---- 1.2 输出目录结构 ----
Private Const CFG_USE_SUBFOLDER As Boolean = True   ' True：输出目录下自动建 STEP\ 与 PDF\ 两个子目录
Private Const CFG_OVERWRITE     As Boolean = True   ' True：覆盖同名文件；False：跳过并在清单里标注

'---- 1.3 工程图查找 ----
Private Const CFG_FIND_DRAWING   As Boolean = True  ' 模型导出后，是否自动找它的工程图一起导出 PDF
Private Const CFG_DRW_ALSO_MODEL As Boolean = True  ' 处理工程图时，是否把它引用的模型也导出 STEP
Private Const CFG_DEEP_SEARCH    As Boolean = True  ' 同名找不到时，扫描目录反查引用关系（慢但准）
Private Const CFG_DEEP_LIMIT     As Long = 1500     ' 深搜最多扫描多少个工程图，防超大目录卡死
Private Const CFG_EXTRA_DIRS     As String = ""     ' 额外搜索目录，多个用 ; 分隔，如 "D:\图纸库"

'---- 1.4 文件夹模式 ----
Private Const CFG_RECURSE_SUBDIR As Boolean = False ' 是否递归子目录

'---- 1.5 其它 ----
Private Const CFG_WRITE_REPORT  As Boolean = True   ' 在输出目录生成 _导出清单.csv
Private Const CFG_USE_PS_DIALOG As Boolean = True   ' 多选文件用 PowerShell 原生对话框
Private Const CFG_QUIET_OPEN    As Boolean = True   ' 静默打开文档（不弹配置 / 参考丢失对话框）

'---- 1.6 上次目录记忆（存于 HKCU\Software\VB and VBA Program Settings\SWauto\Export）----
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
