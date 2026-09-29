
'==========================================================================================
'  第 5 节  入口程序
'    SwAuto_ExportActiveDoc   导出当前打开的文档
'    SwAuto_PickFiles         弹文件对话框，多选文件
'    SwAuto_PickFolder        弹文件夹对话框，批量导出
'==========================================================================================

Public Sub SwAuto_ExportActiveDoc()
    Dim d As Object, sPath As String, col As Collection
    If Not M_Init() Then Exit Sub
    Set d = m_swApp.ActiveDoc
    If d Is Nothing Then
        MsgBox "当前没有打开的文档。", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    sPath = d.GetPathName
    If Len(sPath) = 0 Then
        MsgBox "当前文档还没保存到磁盘，请先保存。", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    Set col = New Collection
    col.Add sPath
    M_Run col, "当前文档：" & U_FileName(sPath)
    M_Cleanup
End Sub

Public Sub SwAuto_PickFiles()
    Dim v As Variant, col As Collection, i As Long
    If Not M_Init() Then Exit Sub
    v = U_PickFiles(Cfg_GetSetting(CFG_KEY_IN))
    If IsEmpty(v) Then
        M_Cleanup
        Exit Sub
    End If
    Set col = New Collection
    For i = LBound(v) To UBound(v)
        If U_IsModel(v(i)) Or U_IsDrawing(v(i)) Then
            If Not F_Contains(col, v(i)) Then col.Add v(i)
        End If
    Next i
    If col.Count = 0 Then
        MsgBox "没有选中可处理的 SolidWorks 文件（.sldprt / .sldasm / .slddrw）。", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    Cfg_PutSetting CFG_KEY_IN, U_Parent(col(1))
    M_Run col, "已选 " & col.Count & " 个文件"
    M_Cleanup
End Sub

Public Sub SwAuto_PickFolder()
    Dim sDir As String, col As Collection
    If Not M_Init() Then Exit Sub
    sDir = U_PickFolder("请选择要导出的源文件夹", Cfg_GetSetting(CFG_KEY_IN))
    If Len(sDir) = 0 Then
        M_Cleanup
        Exit Sub
    End If
    Cfg_PutSetting CFG_KEY_IN, sDir
    Set col = M_ScanFolder(sDir)
    If col.Count = 0 Then
        MsgBox "该文件夹里没有找到 .sldprt / .sldasm / .slddrw 文件。" & vbCrLf & _
               "（如需包含子目录，请把 CFG_RECURSE_SUBDIR 改为 True）", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    M_Run col, "文件夹：" & sDir
    M_Cleanup
End Sub

'---- 5.1 主流程 ----
Private Sub M_Run(ByVal colPaths As Collection, ByVal sDesc As String)
    Dim colModels As Collection, colDrawings As Collection, colDep As Collection
    Dim i As Long, j As Long, sPath As String, sNote As String, sDrw As String

    ' (1) 输出目录：默认上次使用的目录
    m_OutRoot = M_AskOutDir(U_Parent(colPaths(1)))
    If Len(m_OutRoot) = 0 Then Exit Sub              ' 用户取消
    If Not U_MakeDir(m_OutRoot) Then
        MsgBox "无法创建输出目录：" & m_OutRoot, vbCritical, "SWauto"
        Exit Sub
    End If
    If CFG_USE_SUBFOLDER Then
        m_StepDir = U_JoinPath(m_OutRoot, "STEP")
        m_PdfDir = U_JoinPath(m_OutRoot, "PDF")
        U_MakeDir m_StepDir
        U_MakeDir m_PdfDir
    Else
        m_StepDir = m_OutRoot
        m_PdfDir = m_OutRoot
    End If

    ' (2) 分类去重
    Set colModels = New Collection
    Set colDrawings = New Collection
    For i = 1 To colPaths.Count
        sPath = colPaths(i)
        If U_IsModel(sPath) Then
            If Not F_Contains(colModels, sPath) Then colModels.Add sPath
        ElseIf U_IsDrawing(sPath) Then
            If Not F_Contains(colDrawings, sPath) Then colDrawings.Add sPath
        End If
    Next i

    ' (3) 确认
    If MsgBox("来源：" & sDesc & vbCrLf & vbCrLf & _
              "模型：" & colModels.Count & " 个  →  导出 STEP (AP" & CFG_STEP_AP & ")" & vbCrLf & _
              "工程图：" & colDrawings.Count & " 个  →  导出 PDF" & IIf(CFG_FIND_DRAWING, vbCrLf & _
              "（模型还会自动查找对应工程图一起导出）", "") & vbCrLf & vbCrLf & _
              "输出目录：" & m_OutRoot & vbCrLf & vbCrLf & "开始导出？", _
              vbYesNo + vbQuestion, "SWauto 导出确认") <> vbYes Then
        Exit Sub
    End If

    ' (4) 为每个模型查找对应工程图
    If CFG_FIND_DRAWING Then
        For i = 1 To colModels.Count
            sDrw = F_Find(colModels(i), sNote)
            If Len(sDrw) > 0 Then
                If Not F_Contains(colDrawings, sDrw) Then colDrawings.Add sDrw
            Else
                m_NoDrawing.Add colModels(i)
            End If
        Next i
    End If

    ' (5) 工程图反查它引用的模型
    If CFG_DRW_ALSO_MODEL Then
        For i = 1 To colDrawings.Count
            Set colDep = F_ModelsOfDrawing(colDrawings(i))
            For j = 1 To colDep.Count
                If Not F_Contains(colModels, colDep(j)) Then colModels.Add colDep(j)
            Next j
        Next i
    End If

    ' (6) 执行
    For i = 1 To colModels.Count
        X_DoModel colModels(i)
    Next i
    For i = 1 To colDrawings.Count
        X_DoDrawing colDrawings(i)
    Next i

    ' (7) 清单 + 汇总
    If CFG_WRITE_REPORT Then M_WriteReport
    M_RestoreActive
    M_Summary
End Sub

'---- 5.2 输出目录选择 ----
Private Function M_AskOutDir(ByVal sSeed As String) As String
    Dim sDef As String
    sDef = Cfg_GetSetting(CFG_KEY_OUT)
    If Not U_DirExists(sDef) Then sDef = sSeed
    If Not U_DirExists(sDef) Then sDef = ""
    M_AskOutDir = U_PickFolder("请选择导出文件的输出目录", sDef)
    If Len(M_AskOutDir) > 0 Then Cfg_PutSetting CFG_KEY_OUT, M_AskOutDir
End Function

'---- 5.3 扫描文件夹 ----
Private Function M_ScanFolder(ByVal sDir As String) As Collection
    Dim col As New Collection, colDirs As New Collection
    Dim fso As Object, fld As Object, f As Object, sf As Object
    Dim k As Long, s As String
    Set M_ScanFolder = col
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    colDirs.Add U_NormPath(sDir)
    k = 1
    Do While k <= colDirs.Count
        s = colDirs(k)
        k = k + 1
        If fso.FolderExists(s) Then
            Set fld = fso.GetFolder(s)
            For Each f In fld.Files
                If InStr(f.Name, "~$") = 0 Then
                    If U_IsModel(f.Path) Or U_IsDrawing(f.Path) Then col.Add f.Path
                End If
            Next f
            If CFG_RECURSE_SUBDIR Then
                For Each sf In fld.SubFolders
                    colDirs.Add U_NormPath(sf.Path)
                Next sf
            End If
        End If
    Loop
End Function

'---- 5.4 导出清单 ----
Private Sub M_WriteReport()
    Dim sb As String, i As Long, it As Variant
    sb = "SWauto 导出清单" & vbCrLf
    sb = sb & U_Csv("导出时间") & "," & U_Csv(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & vbCrLf
    sb = sb & U_Csv("输出目录") & "," & U_Csv(m_OutRoot) & vbCrLf
    sb = sb & U_Csv("STEP 协议") & "," & U_Csv("AP" & CStr(CFG_STEP_AP)) & vbCrLf
    sb = sb & U_Csv("模型 STEP 成功") & "," & U_CountBy("模型", "成功") & vbCrLf
    sb = sb & U_Csv("模型 STEP 失败") & "," & U_CountBy("模型", "失败") & vbCrLf
    sb = sb & U_Csv("模型 STEP 跳过") & "," & U_CountBy("模型", "跳过") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 成功") & "," & U_CountBy("图纸", "成功") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 失败") & "," & U_CountBy("图纸", "失败") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 跳过") & "," & U_CountBy("图纸", "跳过") & vbCrLf
    sb = sb & U_Csv("未找到工程图的模型") & "," & m_NoDrawing.Count & vbCrLf
    sb = sb & vbCrLf
    sb = sb & "序号,类别,源文件,输出文件,结果,备注" & vbCrLf
    For i = 1 To m_Results.Count
        it = m_Results(i)
        sb = sb & i & "," & U_Csv(it(0)) & "," & U_Csv(it(1)) & "," & _
             U_Csv(it(2)) & "," & U_Csv(it(3)) & "," & U_Csv(it(4)) & vbCrLf
    Next i
    sb = sb & vbCrLf & U_Csv("以下模型未找到对应工程图") & vbCrLf
    If m_NoDrawing.Count = 0 Then
        sb = sb & U_Csv("（无）") & vbCrLf
    Else
        For i = 1 To m_NoDrawing.Count
            sb = sb & U_Csv(m_NoDrawing(i)) & vbCrLf
        Next i
    End If
    U_WriteText U_JoinPath(m_OutRoot, "_导出清单.csv"), sb
End Sub

'---- 5.5 汇总弹窗 ----
Private Sub M_Summary()
    Dim s As String, i As Long, nFail As Long, it As Variant
    s = "导出完成" & vbCrLf & vbCrLf
    s = s & "输出目录：" & m_OutRoot & vbCrLf & vbCrLf
    s = s & "模型 → STEP：" & U_CountBy("模型", "成功") & " 成功"
    If U_CountBy("模型", "失败") > 0 Then s = s & "，" & U_CountBy("模型", "失败") & " 失败"
    If U_CountBy("模型", "跳过") > 0 Then s = s & "，" & U_CountBy("模型", "跳过") & " 跳过"
    s = s & vbCrLf
    s = s & "工程图 → PDF：" & U_CountBy("图纸", "成功") & " 成功"
    If U_CountBy("图纸", "失败") > 0 Then s = s & "，" & U_CountBy("图纸", "失败") & " 失败"
    If U_CountBy("图纸", "跳过") > 0 Then s = s & "，" & U_CountBy("图纸", "跳过") & " 跳过"

    s = s & vbCrLf & vbCrLf
    If m_NoDrawing.Count = 0 Then
        s = s & "所有模型都找到了对应工程图。"
    Else
        s = s & "未找到对应工程图（" & m_NoDrawing.Count & " 个）：" & vbCrLf & _
            U_BriefList(m_NoDrawing, 8)
    End If

    nFail = U_CountBy("模型", "失败") + U_CountBy("图纸", "失败")
    If nFail > 0 Then
        s = s & vbCrLf & vbCrLf & "失败明细：" & vbCrLf
        i = 0
        For Each it In m_Results
            If it(3) = "失败" Then
                i = i + 1
                If i > 5 Then
                    s = s & "…"
                    Exit For
                End If
                s = s & "· " & U_FileName(it(1)) & "：" & it(4) & vbCrLf
            End If
        Next it
    End If

    If CFG_WRITE_REPORT Then
        s = s & vbCrLf & "明细清单：" & U_JoinPath(m_OutRoot, "_导出清单.csv")
    End If
    MsgBox s, IIf(nFail > 0, vbExclamation, vbInformation), "SWauto"
End Sub

'---- 5.6 初始化 / 收尾 ----
Private Function M_Init() As Boolean
    On Error Resume Next
    Set m_swApp = Application.SldWorks
    If m_swApp Is Nothing Then Set m_swApp = CreateObject("SldWorks.Application")
    On Error GoTo 0
    If m_swApp Is Nothing Then
        MsgBox "无法连接 SolidWorks。", vbCritical, "SWauto"
        Exit Function
    End If
    On Error Resume Next
    Set m_UsedNames = CreateObject("Scripting.Dictionary")
    m_UsedNames.CompareMode = 1
    On Error GoTo 0
    Set m_Results = New Collection
    Set m_NoDrawing = New Collection
    Set m_IdxByPath = Nothing
    Set m_IdxByName = Nothing
    m_IdxBuilt = False
    m_OutRoot = ""
    m_StepDir = ""
    m_PdfDir = ""
    m_OriginalActive = ""
    If Not m_swApp.ActiveDoc Is Nothing Then m_OriginalActive = m_swApp.ActiveDoc.GetTitle
    M_Init = True
End Function

Private Sub M_Cleanup()
    Set m_Results = Nothing
    Set m_NoDrawing = Nothing
    Set m_IdxByPath = Nothing
    Set m_IdxByName = Nothing
    Set m_UsedNames = Nothing
End Sub

' 还原用户原来的活动文档，避免批量导出后界面停在最后一个文件上
Private Sub M_RestoreActive()
    Dim lErr As Long
    If Len(m_OriginalActive) = 0 Then Exit Sub
    On Error Resume Next
    m_swApp.ActivateDoc3 m_OriginalActive, False, 0, lErr    ' 0 = 激活时不重建
    On Error GoTo 0
End Sub
