
'==========================================================================================
'  第 2 节  通用工具（路径 / 文件 / 对话框 / 结果记录）
'==========================================================================================

'---- 2.1 路径与文件名 ----
Private Function U_NormPath(ByVal sPath As String) As String
    Dim s As String
    s = Trim$(sPath)
    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    U_NormPath = s
End Function

Private Function U_JoinPath(ByVal sDir As String, ByVal sName As String) As String
    Dim s As String
    s = U_NormPath(sDir)
    If Len(s) = 0 Then
        U_JoinPath = sName
        Exit Function
    End If
    If Right$(s, 1) <> "\" Then s = s & "\"
    U_JoinPath = s & sName
End Function

Private Function U_FileName(ByVal sPath As String) As String
    Dim i As Long
    i = InStrRev(sPath, "\")
    If i = 0 Then U_FileName = sPath Else U_FileName = Mid$(sPath, i + 1)
End Function

Private Function U_Parent(ByVal sPath As String) As String
    Dim i As Long
    i = InStrRev(sPath, "\")
    If i <= 0 Then Exit Function
    If i = 3 Then
        U_Parent = Left$(sPath, 3)          ' 盘根，如 D:\
    Else
        U_Parent = Left$(sPath, i - 1)
    End If
End Function

Private Function U_BaseName(ByVal sPath As String) As String
    Dim s As String, i As Long
    s = U_FileName(sPath)
    i = InStrRev(s, ".")
    If i > 1 Then s = Left$(s, i - 1)
    U_BaseName = s
End Function

Private Function U_Ext(ByVal sPath As String) As String
    Dim s As String, i As Long
    s = U_FileName(sPath)
    i = InStrRev(s, ".")
    If i > 0 Then U_Ext = UCase$(Mid$(s, i)) Else U_Ext = ""
End Function

Private Function U_FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    If Len(sPath) = 0 Then Exit Function
    U_FileExists = (Len(Dir$(sPath, vbNormal)) > 0)
End Function

Private Function U_DirExists(ByVal sPath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    If Len(sPath) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    U_DirExists = fso.FolderExists(sPath)
End Function

' 返回磁盘上的真实大小写路径，便于清单里显示
Private Function U_ResolveCase(ByVal sPath As String) As String
    Dim sDir As String, sName As String, sFound As String
    If Not U_FileExists(sPath) Then Exit Function
    If InStr(sPath, "\") = 0 Then
        U_ResolveCase = sPath
        Exit Function
    End If
    sDir = Left$(sPath, InStrRev(sPath, "\"))
    sName = Mid$(sPath, Len(sDir) + 1)
    sFound = Dir$(sDir & sName, vbNormal)
    If Len(sFound) > 0 Then U_ResolveCase = sDir & sFound Else U_ResolveCase = sPath
End Function

Private Function U_IsModel(ByVal sPath As String) As Boolean
    Dim e As String
    e = U_Ext(sPath)
    U_IsModel = (e = ".SLDPRT" Or e = ".SLDASM")
End Function

Private Function U_IsDrawing(ByVal sPath As String) As Boolean
    U_IsDrawing = (U_Ext(sPath) = ".SLDDRW")
End Function

Private Function U_DocType(ByVal sPath As String) As Long
    Select Case U_Ext(sPath)
        Case ".SLDPRT": U_DocType = swDocumentTypes_e.swDocPART
        Case ".SLDASM": U_DocType = swDocumentTypes_e.swDocASSEMBLY
        Case ".SLDDRW": U_DocType = swDocumentTypes_e.swDocDRAWING
        Case Else:      U_DocType = -1
    End Select
End Function

' 逐级创建目录，支持 UNC；已存在返回 True
Private Function U_MakeDir(ByVal sPath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    sPath = U_NormPath(sPath)
    If Len(sPath) = 0 Then Exit Function
    If fso.FolderExists(sPath) Then
        U_MakeDir = True
        Exit Function
    End If
    If Len(sPath) <= 3 Then Exit Function          ' 连盘根都不在，放弃
    U_MakeDir = U_MakeDir(fso.GetParentFolderName(sPath))
    If U_MakeDir Then fso.CreateFolder sPath
    U_MakeDir = fso.FolderExists(sPath)
End Function

' 给一个源文件算出输出路径；若配置为不覆盖且已存在，返回 "" 表示跳过
Private Function U_TargetPath(ByVal sSrcPath As String, ByVal sExt As String, ByVal sOutDir As String) As String
    Dim sBase As String, sTry As String, i As Long
    sBase = U_BaseName(sSrcPath)
    sTry = U_JoinPath(sOutDir, sBase & sExt)
    If (Not CFG_OVERWRITE) And U_FileExists(sTry) Then
        U_TargetPath = ""
        Exit Function
    End If
    ' 本次运行内的重名（不同目录的同名文件）自动加序号，避免互相覆盖
    If m_UsedNames.Exists(UCase$(sTry)) Then
        For i = 2 To 999
            sTry = U_JoinPath(sOutDir, sBase & "_" & i & sExt)
            If Not m_UsedNames.Exists(UCase$(sTry)) Then Exit For
        Next i
    End If
    m_UsedNames.Add UCase$(sTry), True
    U_TargetPath = sTry
End Function

'---- 2.2 文本 / CSV ----
Private Function U_Csv(ByVal s As String) As String
    U_Csv = """" & Replace(s, """", """""") & """"
End Function

Private Function U_WriteText(ByVal sFile As String, ByVal sText As String) As Boolean
    Dim st As Object, h As Integer
    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    If Not st Is Nothing Then
        st.Type = 2                 ' adTypeText
        st.Charset = "utf-8"
        st.Open
        st.WriteText sText
        st.SaveToFile sFile, 2      ' adSaveCreateOverWrite
        st.Close
    End If
    If Err.Number = 0 And U_FileExists(sFile) Then
        U_WriteText = True
        Exit Function
    End If
    ' 回退：ANSI 直写
    Err.Clear
    h = FreeFile
    Open sFile For Output As #h
    Print #h, sText
    Close #h
    U_WriteText = U_FileExists(sFile)
End Function

Private Function U_ReadLines(ByVal sFile As String) As Variant
    Dim fso As Object, ts As Object, col As Collection, s As String
    Dim a() As String, i As Long
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(sFile, 1, False, 0)
    If ts Is Nothing Then Exit Function
    Set col = New Collection
    Do While Not ts.AtEndOfStream
        s = ts.ReadLine
        If Left$(s, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then s = Mid$(s, 4)   ' 去 BOM
        s = Trim$(s)
        If Len(s) > 0 Then col.Add s
    Loop
    ts.Close
    If col.Count = 0 Then Exit Function
    ReDim a(col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col(i)
    Next i
    U_ReadLines = a
End Function

'---- 2.3 外部进程 ----
' 把字符串包成 PowerShell 单引号字面量（内部单引号翻倍）
Private Function U_PsQuote(ByVal s As String) As String
    U_PsQuote = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function U_ShellRunWait(ByVal sCmd As String, ByVal nStyle As Long) As Long
    Dim wsh As Object
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    U_ShellRunWait = wsh.Run(sCmd, nStyle, True)
End Function

'---- 2.4 文件夹选择框（默认定位到 sDefault，即"上次使用的目录"）----
Private Function U_PickFolder(ByVal sTitle As String, ByVal sDefault As String) As String
    Dim sh As Object, fld As Object
    Dim sOld As String, sRes As String
    On Error Resume Next
    sOld = CurDir
    If U_DirExists(sDefault) Then
        If Mid$(sDefault, 2, 1) = ":" Then        ' UNC 不做 ChDir，会报错
            ChDrive Left$(sDefault, 1)
            ChDir sDefault
        End If
    End If
    Set sh = CreateObject("Shell.Application")
    Set fld = sh.BrowseForFolder(0, sTitle, &H41, 0)   ' 1 = 只返回目录, 64 = 新式对话框（可新建文件夹）
    If Not fld Is Nothing Then
        sRes = fld.Self.Path
        If Len(sRes) = 0 Then sRes = fld.Items().Item().Path    ' 新式对话框下 Self.Path 可能为空
    End If
    If Len(sOld) > 0 Then ChDir sOld
    On Error GoTo 0
    U_PickFolder = U_NormPath(sRes)
End Function

'---- 2.5 文件多选框（PowerShell 原生对话框，支持多选 + 记住初始目录）----
Private Function U_PickFiles(ByVal sDefaultDir As String) As Variant
    Dim sTmp As String, sPs As String, sCmd As String, sInit As String
    Dim v As Variant
    U_PickFiles = Empty
    If Not CFG_USE_PS_DIALOG Then Exit Function

    sTmp = Environ$("TEMP")
    If Len(sTmp) = 0 Then sTmp = Environ$("TMP")
    If Len(sTmp) = 0 Then Exit Function
    sTmp = U_JoinPath(sTmp, "swauto_selection.txt")
    On Error Resume Next
    If Len(Dir$(sTmp, vbNormal)) > 0 Then Kill sTmp
    Err.Clear

    If U_DirExists(sDefaultDir) Then
        sInit = "$d.InitialDirectory = " & U_PsQuote(sDefaultDir) & ";"
    End If
    sPs = "Add-Type -AssemblyName System.Windows.Forms;" & _
          "$o = New-Object System.Windows.Forms.Form;" & _
          "$o.TopMost = $true; $o.ShowInTaskbar = $false; $o.FormBorderStyle = 'None';" & _
          "$o.Width = 1; $o.Height = 1; $o.StartPosition = 'Manual';" & _
          "$o.Left = -3000; $o.Top = -3000; $o.Show();" & _
          "$d = New-Object System.Windows.Forms.OpenFileDialog;" & _
          "$d.Multiselect = $true;" & _
          "$d.Title = '选择要导出的 SolidWorks 文件（可多选）';" & _
          "$d.Filter = 'SolidWorks 文件|*.sldprt;*.sldasm;*.slddrw|所有文件 (*.*)|*.*';" & _
          sInit & _
          "if ($d.ShowDialog($o) -eq [System.Windows.Forms.DialogResult]::OK) {" & _
          "$d.FileNames | Set-Content -LiteralPath " & U_PsQuote(sTmp) & " -Encoding Default };" & _
          "$o.Close()"
    sCmd = "powershell.exe -NoProfile -STA -WindowStyle Hidden -Command """ & sPs & """"
    U_ShellRunWait sCmd, 0

    If Len(Dir$(sTmp, vbNormal)) = 0 Then Exit Function
    v = U_ReadLines(sTmp)
    On Error Resume Next
    Kill sTmp
    U_PickFiles = v
End Function

'---- 2.6 结果记录 ----
Private Sub U_AddResult(ByVal sKind As String, ByVal sSrc As String, _
                       ByVal sOut As String, ByVal sState As String, ByVal sNote As String)
    m_Results.Add Array(sKind, sSrc, sOut, sState, sNote)
End Sub

Private Function U_CountBy(ByVal sKind As String, ByVal sState As String) As Long
    Dim i As Long, it As Variant, n As Long
    For i = 1 To m_Results.Count
        it = m_Results(i)
        If it(0) = sKind And it(3) = sState Then n = n + 1
    Next i
    U_CountBy = n
End Function

Private Function U_BriefList(ByVal col As Collection, ByVal nMax As Long) As String
    Dim i As Long, s As String
    For i = 1 To col.Count
        If i > nMax Then
            s = s & "… 等共 " & col.Count & " 个"
            Exit For
        End If
        If Len(s) > 0 Then s = s & "，"
        s = s & U_FileName(col(i))
    Next i
    U_BriefList = s
End Function
