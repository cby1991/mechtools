
'==========================================================================================
'  第 3.5 节  目标收集（标准件过滤 + 装配体展开）
'            只服务于 SwAuto_ExportActiveDoc：
'            把"当前激活的文档"解析成 一组要导出的模型(→STEP) + 一组工程图(→PDF)
'==========================================================================================

'---- 3.5.1 标准件判定 ----
' 命中任意一条即算标准件：
'   ① SolidWorks 官方 Toolbox 标记（最准，需要拿到 ModelDoc）
'   ② 文件路径里含 CFG_STD_PATH_HINTS
'   ③ 文件名里含 CFG_STD_KEYWORDS / CFG_STD_STDCODES（词边界匹配，不会误伤）
'   ④ 文件夹名含 Toolbox 标准代号（如 ...\ISO\、...\GB\）
'
' 注意：判断"是不是 Toolbox 零件"的正确 API 是
'       ModelDocExtension::ToolboxPartType，
'       而不是 Component2 上的什么 IsToolboxComponent —— 那个属性并不存在（试过，编译不过）。
'       轻化组件的 GetModelDoc2 可能返回 Nothing，此时这一条自动跳过，交给后面的规则兜底。
Private Function G_IsStdPart(ByVal sPath As String, ByVal mDoc As Object, ByRef sReason As String) As Boolean
    Dim i As Long, v As Variant, sName As String
    Dim vTb As Variant
    sReason = ""
    If Not CFG_EXCLUDE_STD_PARTS Then Exit Function

    ' ① 官方 Toolbox 标记
    '    该属性可能返回 Boolean(-1/0)，也可能返回 swToolboxPartType_e 枚举，统一用 <> 0 判断
    If CFG_STD_USE_API And Not mDoc Is Nothing Then
        vTb = Empty
        On Error Resume Next
        vTb = mDoc.Extension.ToolboxPartType
        Err.Clear
        On Error GoTo 0
        If IsNumeric(vTb) Then
            If vTb <> 0 Then
                sReason = "Toolbox 标准件"
                G_IsStdPart = True
                Exit Function
            End If
        End If
    End If

    ' ② 路径关键字
    v = U_KeywordList(CFG_STD_PATH_HINTS)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If InStr(1, sPath, CStr(v(i)), vbTextCompare) > 0 Then
                sReason = "路径含 " & v(i)
                G_IsStdPart = True
                Exit Function
            End If
        Next i
    End If

    ' ③ 文件名关键字
    sName = U_FileName(sPath)
    v = U_KeywordList(CFG_STD_KEYWORDS)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If U_HitKeyword(sName, CStr(v(i))) Then
                sReason = "文件名含 " & v(i)
                G_IsStdPart = True
                Exit Function
            End If
        Next i
    End If

    ' ③ 国标 / 国际标准号
    v = U_KeywordList(CFG_STD_STDCODES)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If U_HitKeyword(sName, CStr(v(i))) Then
                sReason = "标准号 " & v(i)
                G_IsStdPart = True
                Exit Function
            End If
        Next i
    End If
End Function

' 记一个被过滤的标准件（同一次运行内去重），并写进导出清单
Private Sub G_NoteStd(ByVal sPath As String, ByVal sReason As String)
    If F_Contains(m_SkippedStd, sPath) Then Exit Sub
    m_SkippedStd.Add sPath
    U_AddResult "模型", sPath, "", "跳过", "标准件已过滤（" & sReason & "）"
End Sub

' 取组件对应的模型文档；组件被轻化 / 抑制时可能拿到 Nothing（此时官方 Toolbox 标记用不了，
' 后面的路径与文件名规则仍然有效，不影响整体判定）
Private Function G_DocOfComp(ByVal comp As Object) As Object
    If comp Is Nothing Then Exit Function
    On Error Resume Next
    Set G_DocOfComp = comp.GetModelDoc2
    Err.Clear
    On Error GoTo 0
End Function

Private Sub G_AddModel(ByVal sPath As String, ByVal colModels As Collection, ByRef n As Long)
    If Not F_Contains(colModels, sPath) Then
        colModels.Add sPath
        n = n + 1
    End If
End Sub

'---- 3.5.2 装配体展开 ----
' 遍历装配体（含各级子装配体）把零件收进 colModels；标准件只记不打
Private Function G_ExpandAssembly(ByVal d As Object, ByVal sAsmPath As String, _
                                  ByVal colModels As Collection) As Long
    Dim vComps As Variant, i As Long, comp As Object, sP As String, sR As String, nPart As Long
    Dim mdoc As Object

    ' 轻化组件先还原，否则 GetComponents 可能漏件、也取不到路径
    On Error Resume Next
    d.ResolveAllLightWeightComponents True
    Err.Clear
    On Error GoTo 0

    ' False = 展开所有层级（含子装配体内部的组件）
    On Error Resume Next
    vComps = d.GetComponents(False)
    Err.Clear
    On Error GoTo 0

    If IsArray(vComps) Then
        For i = LBound(vComps) To UBound(vComps)
            Set comp = Nothing
            sP = ""
            On Error Resume Next
            Set comp = vComps(i)
            If Not comp Is Nothing Then sP = comp.GetPathName
            Err.Clear
            On Error GoTo 0

            If Len(sP) > 0 Then
                Select Case U_Ext(sP)
                Case ".SLDPRT"
                    Set mdoc = G_DocOfComp(comp)
                    If G_IsStdPart(sP, mdoc, sR) Then
                        G_NoteStd sP, sR
                    Else
                        G_AddModel sP, colModels, nPart
                    End If
                Case ".SLDASM"
                    ' 子装配体自身是否也出一份 STEP
                    If CFG_ASM_SELF Then
                        If StrComp(sP, sAsmPath, vbTextCompare) <> 0 Then G_AddModel sP, colModels, nPart
                    End If
                End Select
            Else
                ' 虚拟零件（几何存在装配体内部）没法单独导出成 STEP，记一笔免得用户以为漏了
                On Error Resume Next
                If Not comp Is Nothing Then sP = comp.Name2
                Err.Clear
                On Error GoTo 0
                If Len(sP) > 0 Then
                    U_AddResult "模型", sP, "", "跳过", "虚拟零件（保存在装配体内部，无法单独导出）"
                End If
            End If
        Next i
    End If

    If nPart = 0 Then G_DepsOfAssembly sAsmPath, colModels   ' 兜底：直接读文件依赖表
    G_ExpandAssembly = nPart
End Function

' 兜底通道：装配体打开异常 / GetComponents 空返时，读文件依赖表（不需要打开文档）
Private Function G_DepsOfAssembly(ByVal sAsmPath As String, ByVal colModels As Collection) As Long
    Dim vDeps As Variant, i As Long, sItem As String, n As Long, sR As String
    On Error Resume Next
    vDeps = m_swApp.GetDocumentDependencies2(sAsmPath, True, False, False)
    Err.Clear
    On Error GoTo 0
    If Not IsArray(vDeps) Then Exit Function
    For i = LBound(vDeps) To UBound(vDeps)
        sItem = CStr(vDeps(i))
        If Len(sItem) > 0 And InStr(sItem, "\") > 0 Then
            Select Case U_Ext(sItem)
            Case ".SLDPRT"
                If G_IsStdPart(sItem, Nothing, sR) Then
                    G_NoteStd sItem, sR
                Else
                    G_AddModel sItem, colModels, n
                End If
            Case ".SLDASM"
                If CFG_ASM_SELF Then
                    If StrComp(sItem, sAsmPath, vbTextCompare) <> 0 Then G_AddModel sItem, colModels, n
                End If
            End Select
        End If
    Next i
    G_DepsOfAssembly = n
End Function

'---- 3.5.3 入口：按"当前激活的文档"收集 ----
' 返回 False 表示无法处理或用户取消
Private Function G_CollectActive(ByRef colModels As Collection, _
                                 ByRef colDrawings As Collection, _
                                 ByRef sDesc As String, ByRef sExtra As String) As Boolean
    Dim d As Object, sPath As String, sR As String, n As Long, t As Long

    Set colModels = New Collection
    Set colDrawings = New Collection
    sDesc = ""
    sExtra = ""

    Set d = m_swApp.ActiveDoc
    If d Is Nothing Then
        MsgBox "当前没有打开的文档。" & vbCrLf & vbCrLf & _
               "这个宏要读「当前激活的那个窗口」，请先打开一个零件 / 装配体 / 工程图；" & vbCrLf & _
               "想手动挑文件的话，请改用 SwAuto_PickFiles。", vbExclamation, "SWauto"
        Exit Function
    End If
    sPath = d.GetPathName
    If Len(sPath) = 0 Then
        MsgBox "当前文档还没有保存到磁盘，请先保存后再运行。", vbExclamation, "SWauto"
        Exit Function
    End If

    t = d.GetType
    Select Case t

    Case swDocumentTypes_e.swDocPART
        sDesc = "零件：" & U_FileName(sPath)
        If G_IsStdPart(sPath, d, sR) Then
            G_NoteStd sPath, sR
            sExtra = "该零件被判定为标准件（" & sR & "），已跳过。" & vbCrLf & _
                     "如果它其实需要导出，请到配置第 1.6 节删掉对应关键字。"
        Else
            colModels.Add sPath
            sExtra = "本零件 → STEP；它对应的工程图 → PDF。"
        End If

    Case swDocumentTypes_e.swDocASSEMBLY
        sDesc = "装配体：" & U_FileName(sPath)
        If CFG_ASM_EXPAND Then
            If CFG_ASM_SELF Then G_AddModel sPath, colModels, n
            n = G_ExpandAssembly(d, sPath, colModels)
            sExtra = "装配体已展开为 " & colModels.Count & " 个模型，每个都会导出 STEP 并查找对应工程图。"
        Else
            colModels.Add sPath
            sExtra = "装配体未展开（配置 CFG_ASM_EXPAND = False），只导出装配体本身。"
        End If

    Case swDocumentTypes_e.swDocDRAWING
        sDesc = "工程图：" & U_FileName(sPath)
        colDrawings.Add sPath
        sExtra = "本工程图 → PDF；它引用的模型会一并导出 STEP。"

    Case Else
        MsgBox "当前文档不是 SolidWorks 零件 / 装配体 / 工程图，无法处理。", vbExclamation, "SWauto"
        Exit Function

    End Select

    ' 确认框里标清楚是哪种模式，避免用户点错宏时看不出来
    sDesc = "【当前激活文档】" & sDesc
    G_CollectActive = True
End Function
