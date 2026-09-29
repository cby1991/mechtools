
'==========================================================================================
'  第 4 节  导出核心（STEP / PDF）
'==========================================================================================

'---- 4.1 打开 / 关闭文档 ----
Private Function X_FindOpenDoc(ByVal sPath As String) As Object
    Dim vDocs As Variant, i As Long, d As Object
    On Error Resume Next
    vDocs = m_swApp.GetDocuments
    If IsArray(vDocs) Then
        For i = LBound(vDocs) To UBound(vDocs)
            Set d = vDocs(i)
            If Not d Is Nothing Then
                If StrComp(d.GetPathName, sPath, vbTextCompare) = 0 Then
                    Set X_FindOpenDoc = d
                    Exit Function
                End If
            End If
        Next i
    End If
    On Error GoTo 0
End Function

' 已在内存里的文档直接复用；否则静默打开，bOurs=True 表示是我们打开的（跑完要关掉）
Private Function X_Open(ByVal sPath As String, ByVal nType As Long, ByRef bOurs As Boolean) As Object
    Dim d As Object, lErr As Long, lWarn As Long
    bOurs = False
    Set d = X_FindOpenDoc(sPath)
    If Not d Is Nothing Then
        Set X_Open = d
        Exit Function
    End If
    On Error Resume Next
    If CFG_QUIET_OPEN Then
        Set d = m_swApp.OpenDoc6(sPath, nType, swOpenDocOptions_e.swOpenDocOptions_Silent, "", lErr, lWarn)
    Else
        Set d = m_swApp.OpenDoc6(sPath, nType, 0, "", lErr, lWarn)
    End If
    If Not d Is Nothing Then bOurs = True
    Set X_Open = d
End Function

Private Sub X_Close(ByVal sPath As String, ByVal bOurs As Boolean)
    Dim d As Object
    If Not bOurs Then Exit Sub
    On Error Resume Next
    Set d = X_FindOpenDoc(sPath)
    If Not d Is Nothing Then m_swApp.CloseDoc d.GetTitle
End Sub

'---- 4.2 模型 -> STEP ----
Private Function X_ToStep(ByVal d As Object, ByVal sOut As String, ByRef sMsg As String) As Boolean
    Dim ext As Object, lErr As Long, lWarn As Long, bRet As Boolean, sNote As String
    On Error GoTo EH

    ' STEP 协议版本：swStepAP 直接接受 203 / 214
    m_swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swStepAP, CFG_STEP_AP
    If m_swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP) <> CFG_STEP_AP Then
        sNote = "AP" & CFG_STEP_AP & " 未能写入用户选项，已按当前选项导出"
    End If

    ' 装配体先把轻化零件还原，否则可能导出空几何
    If d.GetType = swDocumentTypes_e.swDocASSEMBLY Then
        On Error Resume Next
        d.ResolveAllLightWeightComponents True
        Err.Clear
        On Error GoTo EH
    End If

    Set ext = d.Extension
    bRet = ext.SaveAs(sOut, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                      swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, lErr, lWarn)
    If bRet And lErr = 0 Then
        X_ToStep = True
        sMsg = sNote
    Else
        sMsg = "SaveAs 失败（err=" & lErr & " / warn=" & lWarn & "）"
    End If
    Exit Function
EH:
    sMsg = "异常 " & Err.Number & ": " & Err.Description
End Function

'---- 4.3 工程图 -> PDF ----
Private Function X_ToPdf(ByVal d As Object, ByVal sOut As String, ByRef sMsg As String) As Boolean
    Dim pdfData As Object, lErr As Long, lWarn As Long, bRet As Boolean
    On Error GoTo EH
    Set pdfData = m_swApp.GetExportFileData(swExportDataFileType_e.swExportPDFData)
    If pdfData Is Nothing Then
        sMsg = "无法获取 PDF 导出设置对象"
        Exit Function
    End If
    On Error Resume Next
    pdfData.ExportAsSinglePDF = True        ' 多张图纸合并成一个 PDF
    pdfData.ViewPdfAfterSaving = False      ' 导出后不自动打开阅读器
    Err.Clear
    On Error GoTo EH

    bRet = d.Extension.SaveAs(sOut, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                              swSaveAsOptions_e.swSaveAsOptions_Silent, pdfData, lErr, lWarn)
    If bRet And lErr = 0 Then
        X_ToPdf = True
    Else
        sMsg = "SaveAs 失败（err=" & lErr & " / warn=" & lWarn & "）"
    End If
    Exit Function
EH:
    sMsg = "异常 " & Err.Number & ": " & Err.Description
End Function

'---- 4.4 处理一个模型文件：导出 STEP ----
Private Sub X_DoModel(ByVal sPath As String)
    Dim d As Object, bOurs As Boolean, sOut As String, sMsg As String
    If Not U_FileExists(sPath) Then
        U_AddResult "模型", sPath, "", "失败", "源文件不存在"
        Exit Sub
    End If
    sOut = U_TargetPath(sPath, CFG_STEP_EXT, m_StepDir)
    If Len(sOut) = 0 Then
        U_AddResult "模型", sPath, U_JoinPath(m_StepDir, U_BaseName(sPath) & CFG_STEP_EXT), "跳过", "输出文件已存在"
        Exit Sub
    End If
    Set d = X_Open(sPath, U_DocType(sPath), bOurs)
    If d Is Nothing Then
        U_AddResult "模型", sPath, sOut, "失败", "无法打开文档（请检查参考引用是否齐全）"
        Exit Sub
    End If
    If X_ToStep(d, sOut, sMsg) Then
        U_AddResult "模型", sPath, sOut, "成功", sMsg
    Else
        U_AddResult "模型", sPath, sOut, "失败", sMsg
    End If
    X_Close sPath, bOurs
End Sub

'---- 4.5 处理一个工程图文件：导出 PDF ----
Private Sub X_DoDrawing(ByVal sPath As String)
    Dim d As Object, bOurs As Boolean, sOut As String, sMsg As String
    If Not U_FileExists(sPath) Then
        U_AddResult "图纸", sPath, "", "失败", "源文件不存在"
        Exit Sub
    End If
    sOut = U_TargetPath(sPath, CFG_PDF_EXT, m_PdfDir)
    If Len(sOut) = 0 Then
        U_AddResult "图纸", sPath, U_JoinPath(m_PdfDir, U_BaseName(sPath) & CFG_PDF_EXT), "跳过", "输出文件已存在"
        Exit Sub
    End If
    Set d = X_Open(sPath, swDocumentTypes_e.swDocDRAWING, bOurs)
    If d Is Nothing Then
        U_AddResult "图纸", sPath, sOut, "失败", "无法打开工程图（请检查其引用的模型是否可访问）"
        Exit Sub
    End If
    If X_ToPdf(d, sOut, sMsg) Then
        U_AddResult "图纸", sPath, sOut, "成功", sMsg
    Else
        U_AddResult "图纸", sPath, sOut, "失败", sMsg
    End If
    X_Close sPath, bOurs
End Sub
