Attribute VB_Name = "CodeModuleUtil"
Option Explicit


'https://www.excel-ubara.com/excelvba5/EXCELVBA269.html


'Dictionary�Ƀv���V�[�W���[�E�v���p�e�B�����i�[
Public Sub getCodeModule(ByRef dicProcInfo As Dictionary, _
                         ByVal wb As Workbook, _
                         ByVal sMod As String, _
                         ByVal ModuleType As String)
    Dim cProcInfo As ProcedureInfo
    Dim sProcName As String
    Dim sProcKey As String
    Dim iProcKind As Long: iProcKind = 0
    Dim i As Long
    Dim VVC As Object
    Set VVC = wb.VBProject.VBComponents(sMod).CodeModule
    i = 1
    Do While i <= VVC.CountOfLines
        sProcName = VVC.ProcOfLine(i, iProcKind)
        sProcKey = sMod & "." & sProcName
        If sProcName <> "" Then
            If isProcLine(VVC.Lines(i, 1), sProcName) Then
                If Not dicProcInfo.Exists(sProcKey) Then
                    Set cProcInfo = New ProcedureInfo
                    cProcInfo.ModName = sMod
                    cProcInfo.ModType = ModuleType
                    cProcInfo.ProcName = sProcName
                    cProcInfo.ProcKind = iProcKind
                    cProcInfo.LineNo = i
                    cProcInfo.Comment = getProcComment(i, VVC)
                    cProcInfo.Source = getProcSource(i, VVC)
                    dicProcInfo.Add sProcKey, cProcInfo
                End If
            End If
        End If
        i = i + 1
    Loop
End Sub

'�v���V�[�W���[�E�v���p�e�B��`�s���̔���
Private Function isProcLine(ByVal strLine As String, _
                            ByVal ProcName As String) As Boolean
    strLine = " " & Trim(strLine)
    Select Case True
        Case Left(strLine, 1) = " '"
            isProcLine = False
        Case strLine Like "* Sub " & ProcName & "(*"
            isProcLine = True
        Case strLine Like "* Sub " & ProcName & " _"
            isProcLine = True
        Case strLine Like "* Function " & ProcName & "(*"
            isProcLine = True
        Case strLine Like "* Function " & ProcName & " _"
            isProcLine = True
        Case strLine Like "* Property * " & ProcName & "(*"
            isProcLine = True
        Case strLine Like "* Property * " & ProcName & " _"
            isProcLine = True
        Case Else
            isProcLine = False
    End Select
End Function

'�p���s( _)�S�Ă�A������������ŕԂ�
Private Function getProcSource(ByRef i As Long, _
                               ByVal aCodeModule As Object) As String
    getProcSource = ""
    Dim sTemp As String
    Do
        sTemp = Trim(aCodeModule.Lines(i, 1))
        If Right(aCodeModule.Lines(i, 1), 2) = " _" Then
            sTemp = Left(sTemp, Len(sTemp) - 1)
        End If
        getProcSource = getProcSource & sTemp
        If Right(aCodeModule.Lines(i, 1), 2) <> " _" Then Exit Do
        i = i + 1
    Loop
End Function

'�v���V�[�W���[�̒��O�̃R�����g���擾
Private Function getProcComment(ByVal i As Long, _
                                ByVal aCodeModule As Object) As String
    getProcComment = ""
    i = i - 1
    Do While Left(aCodeModule.Lines(i, 1), 1) = "'"
        If getProcComment <> "" Then getProcComment = vbLf & getProcComment
        getProcComment = aCodeModule.Lines(i, 1) & getProcComment
        i = i - 1
    Loop
End Function

