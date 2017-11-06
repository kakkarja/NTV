Attribute VB_Name = "NTV_Bas_Mod"
Option Explicit

'''Workbook event Codes'''
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim npath As String
Dim cName As String
Dim conf As Variant
    Application.DisplayAlerts = False
    If Right(ThisWorkbook.Name, 4) = _
    "xltm" Then
        ThisWorkbook.SaveAs _
    Environ("userprofile") & "\Documents\" _
    & ThisWorkbook.WriteReservedBy _
    & Format(Now, "DDMMYY_hhmmss") _
    , xlOpenXMLWorkbookMacroEnabled
    Else
Oboy:
        On Error GoTo 0
        conf = MsgBox( _
        "Save with different name?", _
        vbYesNo, "Note Tree View")
        Select Case conf
            Case Is = vbYes
                On Error Resume Next
                cName = InputBox( _
                "Name your file before save:", _
                "Note Tree View")
                ThisWorkbook.SaveAs _
                Environ("userprofile") & "\Documents\" _
                & cName, xlOpenXMLWorkbookMacroEnabled
                If Err.Number <> 0 Then
                    MsgBox "If you do not want to save" _
                    & " to new file, please choose NO!" _
                    , vbInformation, "Note Tree View"
                    GoTo Oboy
                End If
            Case Else
                ThisWorkbook.SaveAs _
                ThisWorkbook.Path _
                & ThisWorkbook.Name, xlOpenXMLWorkbookMacroEnabled
        End Select
    End If
End Sub

'''Password setup for locking file Userform'''
Private Sub SetP_Click()
    If Pssd = "" Then
        MsgBox "Please submit your password", , _
        "Password Setup"
        Exit Sub
    End If
    With ActiveWorkbook
        If .Path = "" Then
            MsgBox "Please save the workbook first," & _
            " in order to setup a password.", , _
            "Password Setup"
            Unload Me
        Else
            Application.DisplayAlerts = False
            .SaveAs .Path & "\" & .Name, , Pssd
            MsgBox "Please do not forget your password." _
            & " You have just secured your workbook" _
            & " viewing.", vbInformation, _
            "Password Setup"
            Application.DisplayAlerts = True
            Unload Me
        End If
    End With
End Sub

'''For Formatting Userform'''
Private Sub Bol_Click()
    With ActiveCell
        With .Offset(, 1)
        If .Font.Bold = False Then
            .Font.Bold = True
        Else
            .Font.Bold = False
        End If
        End With
    End With
End Sub

Private Sub Ital_Click()
    With ActiveCell
        With .Offset(, 1)
        If .Font.Italic = False Then
            .Font.Italic = True
        Else
            .Font.Italic = False
        End If
        End With
    End With
End Sub

Private Sub Lowe_Click()
    With ActiveCell
        With .Offset(, 1)
            If TypeName(.Value) = "String" Then
                .Value = StrConv(.Value, vbLowerCase)
            End If
        End With
    End With
End Sub

Private Sub Prop_Click()
    With ActiveCell
        With .Offset(, 1)
            If TypeName(.Value) = "String" Then
                .Value = StrConv(.Value, vbProperCase)
            End If
        End With
    End With

End Sub

Private Sub ScrollBar1_Change()
Dim Rn As Range
    With ScrollBar1
        Set Rn = Cells(.Value, 1).EntireRow.Find(Chr(149))
        If Rn.EntireRow.Hidden = False Then
            Rn.Select
        End If
    End With
Set Rn = Nothing
End Sub

Private Sub Stri_Click()
    With ActiveCell
        With .Offset(, 1)
        If .Font.Strikethrough = False Then
            .Font.Strikethrough = True
        Else
            .Font.Strikethrough = False
        End If
        End With
    End With

End Sub

Private Sub Unde_Click()
    With ActiveCell
        With .Offset(, 1)
        If .Font.Underline = -4142 Then
            .Font.Underline = 2
        ElseIf .Font.Underline = 2 Then
            .Font.Underline = -4119
        Else
            .Font.Underline = -4142
        End If
        End With
    End With
End Sub

Private Sub Upp_Click()
    With ActiveCell
        With .Offset(, 1)
        If TypeName(.Value) = "String" Then
            .Value = StrConv(.Value, vbUpperCase)
        End If
        End With
    End With
End Sub
Private Sub Lett_Click()
    With ActiveCell
        With .Offset(, 1)
            On Error Resume Next
            .Font.Color = LiCol.Value
        End With
        .Font.Color = LiCol.Value
    End With
End Sub
Private Sub LiCol_Change()
    With LiCol
        .ForeColor = .Value
        .BackColor = .Value
    End With
End Sub
Private Sub Clea_Click()
    With ActiveCell
        If .Interior.ColorIndex <> -4142 Then
            .Interior.ColorIndex = -4142
        End If
        If .Font.Color <> False Then
            .Font.Color = False
            .Offset(, 1).Font.Color = False
        End If
    End With
End Sub

Private Sub Inte_Click()
    With ActiveCell
        On Error Resume Next
        .Interior.Color = LiCol.Value
    End With
End Sub

Private Sub UserForm_Initialize()
Dim N As Long
Dim NCol As Integer
Dim Marr As Variant
Marr = Array(0, 16777215, 255 _
, 65280, 16711680, 65535, 16711935 _
, 16776960, 128, 32768, 8388608, 32896 _
, 8388736, 8421376, 12632256, 8421504 _
, 16751001, 6697881, 13434879, 16777164 _
, 6684774, 8421631, 13395456, 16764108 _
, 16763904, 13434828, 10092543, 16764057 _
, 13408767, 16751052, 10079487, 16737843 _
, 13421619, 52377, 52479, 39423, 26367 _
, 10053222, 9868950, 6697728, 6723891 _
, 13056, 13107, 13209, 10040115, 3355443)
    For NCol = 0 To UBound(Marr)
        With LiCol
            .AddItem Marr(NCol)
            .ForeColor = Marr(NCol)
            .BackColor = Marr(NCol)
        End With
    Next NCol

    N = Cells.SpecialCells(xlCellTypeConstants).Count / 2
    With ScrollBar1
        .Value = RaRo
        .MIn = 1
        .Max = N
    End With
N = 0
NCol = 0
Marr = Null
End Sub

Private Sub UserForm_Terminate()
    With ActiveCell
        If .Offset(, 1).Font.Strikethrough _
        = True Then
            NoteTV.CheckBox1 = True
        Else
            NoteTV.CheckBox1 = False
        End If
        NoteTV.TextBox1 = .Offset(, 1).Value
    End With
End Sub

'''Note Tree View Userform'''
Private Sub AdR_Click()
    With ActiveCell
        If .Offset(, 1) = "" Then GoTo Bye
        .EntireRow.Insert xlShiftDown, True
    End With
    AdBut
    ChAp
Bye:
End Sub
Private Sub AdBut()
    With ActiveCell
        .EntireRow.Clear
        .Value = Chr(149)
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
End Sub
Private Sub CalAd_Click()
Dim ToN As Long
Dim Rn As Range, RnS As Range
Fx1
    On Error Resume Next
    ToN = Cells.SpecialCells(xlCellTypeConstants).Count
    If Err.Number <> 0 Then
        With Cells(1, 1)
            .Select
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Value = Chr(149)
        End With
    End If
    If ActiveCell.Offset(, 1) = "" Then GoTo Bye
    If ToN Mod 2 = 0 Then
        For Each Rn In Cells(ToN / 2, 1).EntireRow
            If Rn = Chr(149) Then
                Set RnS = Rn.Find(Chr(149))
                With RnS.Offset(1)
                    .Select
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .Value = Chr(149)
                End With
                Set RnS = Nothing
                Exit For
            End If
        Next Rn
    Else
        For Each Rn In Cells((ToN - 1) / 2, 1).EntireRow
            If Rn = Chr(149) Then
                Set RnS = Rn.Find(Chr(149))
                RnS.Offset(1).Select
                Set RnS = Nothing
                Exit For
            End If
        Next Rn
    End If
Bye:
ChAp
ToN = 0
Set Rn = Nothing
Set RnS = Nothing
Fx2
End Sub


Private Sub CheckBox1_Click()
Dim N As Integer
Dim Rn As Range
    If CheckBox1 = True Then
        With ActiveCell
            .Offset(, 1).Font.Strikethrough = True
            N = 1
            Do Until .Offset(N) <> ""
                Set Rn = Nothing
                Set Rn = .Offset(N).EntireRow.Find(Chr(149))
                On Error GoTo Bye
                If Rn.Column > .Column Then
                With Rn
                    .Offset(, 1).Font.Strikethrough = True
                End With
                End If
                N = N + 1
            Loop
        End With
    Else
        With ActiveCell
            .Offset(0, 1).Font.Strikethrough = False
            N = 1
            Do Until .Offset(N) <> ""
                Set Rn = Nothing
                Set Rn = .Offset(N).EntireRow.Find(Chr(149))
                On Error GoTo Bye
                If Rn.Column > .Column Then
                With Rn
                    .Offset(, 1).Font.Strikethrough = False
                End With
                End If
                N = N + 1
            Loop
        End With
    End If
Bye:
N = 0
Set Rn = Nothing
End Sub

Private Sub Dele_Click()
Dim N As Integer
Dim Rn As Range
Fx1
    With ActiveCell
        If .Offset(, 1).Font _
        .Strikethrough = False _
        Or .Offset(1).EntireRow _
        .Hidden = True Then GoTo Bye
        N = 1
        Do Until .Offset(N) <> ""
            Set Rn = Nothing
            Set Rn = .Offset(N).EntireRow.Find(Chr(149))
            On Error GoTo Nex
            If Rn.Column > .Column Then
            With Rn
                If .Offset(, 1) _
                .Font.Strikethrough = True Then
                    N = N + 1
                Else
                    Exit Do
                End If
            End With
            End If
        Loop
    End With
Nex:
    If N = 0 Then
        GoTo Bye
    Else
        ActiveCell.Resize(N).EntireRow.Delete
    End If
    LBut
Bye:
ChAp
N = 0
Set Rn = Nothing
Fx2
End Sub
Private Sub LBut()
Dim LB As Long
Dim Rn As Range
    If ActiveCell = "" Then
        With ActiveCell
            Set Rn = .EntireRow.Find(Chr(149))
            If Not Rn Is Nothing Then
                Rn.Select
            Else
                On Error GoTo Bye
                LB = Cells.SpecialCells _
                (xlCellTypeConstants).Count / 2
                Set Rn = Nothing
                Set Rn = Cells(LB, 1).EntireRow.Find(Chr(149))
                Rn.Select
            End If
        End With
    ElseIf Len(ActiveCell) > 1 Then
        ActiveCell.Offset(, -1).Select
    End If
Bye:
LB = 0
Set Rn = Nothing
End Sub

Private Sub Formatt_Click()
    On Error GoTo Bye
    FormatO.Show
Bye:
End Sub

Private Sub HidRs_Click()
Dim ToRs As Long, N As Integer
Dim Rn As Range
Fx1
    On Error GoTo Bye
    ToRs = Cells.SpecialCells(xlCellTypeConstants).Count / 2
    With ActiveCell
        N = 1
        Do Until .Offset(N) <> ""
            With .Offset(N)
                Set Rn = Nothing
                Set Rn = .EntireRow.Find(Chr(149))
                With Rn
                On Error GoTo Nex
                If .Column >= _
                ActiveCell.Column Then
                    N = N + 1
                Else
                    Exit Do
                End If
                End With
            End With
        Loop
Nex:
        If N = 1 Then
            GoTo Bye
        Else
            If .Offset(1).EntireRow.Hidden = False Then
            .Offset(1).Resize(N - 1).EntireRow.Hidden = True
            End If
        End If
    End With
Bye:
ToRs = 0
N = 0
Set Rn = Nothing
Fx2
End Sub

Private Sub MLe_Click()
    With ActiveCell
        If .Value = Chr(149) Then
            On Error GoTo Bye
            .Resize(, 2).Cut .Offset(, -1)
            .Select
        End If
    End With
    ReF
Bye:
End Sub

Private Sub ReF()
    On Error Resume Next
    With ActiveCell
        If .Offset(-1) <> Chr(149) Then
            .Offset(-1).Font.Bold = True
        ElseIf .Offset(-1) = Chr(149) Then
            .Offset(-1, 1).Font.Bold = False
        End If
        
        If .Offset(1) = Chr(149) Then
            .Offset(, 1).Font.Bold = False
        ElseIf .Offset(1, 1) = Chr(149) Then
            .Offset(, 1).Font.Bold = True
        End If
        
    End With
End Sub

Private Sub MRi_Click()
    With ActiveCell
        If .Value = Chr(149) Then
            .Resize(, 2).Cut .Offset(, 1)
            .Select
        End If
    End With
    ReF
End Sub


Private Sub Sea_Click()
Dim UI As Long
On Error GoTo Bye
UI = Cells.SpecialCells(xlCellTypeConstants).Count
Search.Show
Bye:
UI = 0
End Sub

Private Sub SpinButton1_SpinDown()
Dim Rn As Range
Dim N As Integer
With ActiveCell
'    If .Offset(, 1) = "" Then GoTo Bye
    Set Rn = .Offset(1).EntireRow.Find(Chr(149))
    If Not Rn Is Nothing Then
        With Rn
            Do Until .Offset(N).EntireRow.Hidden = False
                N = N + 1
            Loop
        End With
    End If
End With
    On Error GoTo Bye
    If N = 0 Then
        Rn.Select
    Else
        Set Rn = Rn.Offset(N).EntireRow.Find(Chr(149))
        Rn.Select
    End If
Bye:
    ChAp
Set Rn = Nothing
N = 0
End Sub

Private Sub ChAp()
    With ActiveCell
        If .Offset(, 1).Font.Strikethrough _
        = True Then
            CheckBox1 = True
        Else
            CheckBox1 = False
        End If
        TextBox1 = .Offset(, 1).Value
    End With
        
End Sub

Private Sub SpinButton1_SpinUp()
Dim Rn As Range
Dim N As Integer
With ActiveCell
'    If .Offset(, 1) = "" Then GoTo Bye
    On Error GoTo Bye
    Set Rn = .Offset(-1).EntireRow.Find(Chr(149))
    If Not Rn Is Nothing Then
        With Rn
            Do Until .Offset(-N).EntireRow.Hidden = False
                N = N + 1
            Loop
        End With
    End If
End With
    On Error GoTo Bye
    If N = 0 Then
        Rn.Select
    Else
        Set Rn = Rn.Offset(-N).EntireRow.Find(Chr(149))
        Rn.Select
    End If
Bye:
ChAp
Set Rn = Nothing

End Sub

Private Sub STi_Click()
    With ActiveCell
        If .Comment Is Nothing Then
            If .Value = Chr(149) Then
            .AddComment Format(Now, "DD/MM/YY hh:mm:ss")
            .Comment.Shape.TextFrame.AutoSize = True
            End If
        Else
            .Comment.Delete
        End If
    End With
End Sub

Private Sub TextBox1_Change()
    With ActiveCell
        If .Value = Chr(149) Then
            .Offset(, 1) = TextBox1.Text
        End If
    End With
End Sub

Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBox1.SetFocus
End Sub

Private Sub UHidRs_Click()
Dim ToRs As Long, N As Integer
Dim Rn As Range
Fx1
    On Error GoTo Bye
    ToRs = Cells.SpecialCells(xlCellTypeConstants).Count / 2
    With ActiveCell
        N = 1
        Do Until .Offset(N) <> ""
            With .Offset(N)
                Set Rn = Nothing
                Set Rn = .EntireRow.Find(Chr(149))
                With Rn
                On Error GoTo Nex
                If .Column >= _
                ActiveCell.Column Then
                    N = N + 1
                Else
                    Exit Do
                End If
                End With
            End With
        Loop
Nex:
        If N = 1 Then
            GoTo Bye
        Else
            If .Offset(1).EntireRow.Hidden = True Then
            .Offset(1).Resize(N - 1).EntireRow.Hidden = False
            End If
        End If
    End With
Bye:
ToRs = 0
N = 0
Set Rn = Nothing
Fx2
End Sub

Private Sub UserForm_Initialize()
ActiveSheet.Unprotect Environ("userprofile")
Cells.ColumnWidth = WiAd
ActiveWindow.DisplayGridlines = False
ChAp
End Sub

Private Sub UserForm_Terminate()
    With ActiveSheet
        .Protect Environ("userprofile")
        .EnableSelection = xlNoSelection
    End With
End Sub

Private Sub Fast(SU As Boolean, DS As Boolean, C As String, EE As Boolean)
    Application.ScreenUpdating = SU
    Application.DisplayStatusBar = DS
    Application.Calculation = C
    Application.EnableEvents = EE
End Sub
Private Sub Fx1()
Call Fast(False, True, xlCalculationManual, False)
End Sub
Private Sub Fx2()
Call Fast(True, True, xlCalculationAutomatic, True)
End Sub

'''Search Userform'''
Private Sub Search_Click()
Dim Rn As Range
On Error GoTo Bye
Set Rn = Cells(TextS, 1).EntireRow.Find(Chr(149))
    If Rn.EntireRow.Hidden = False Then
        Rn.Select
        Unload Me
    End If
Bye:
Set Rn = Nothing
End Sub

Private Sub Search_Enter()
Search_Click
End Sub

Private Sub TextS_Change()
Dim AlB As Long
    AlB = Cells.SpecialCells(xlCellTypeConstants).Count / 2
    If IsNumeric(TextS) = False Then
        TextS = ""
    ElseIf TextS > AlB Then
        TextS = ""
    End If
AlB = 0
End Sub

Private Sub TextS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextS.SetFocus
End Sub

Private Sub UserForm_Initialize()
    TextS = fLR
End Sub

Private Sub UserForm_Terminate()
    With ActiveCell
        If .Offset(, 1).Font.Strikethrough _
        = True Then
            NoteTV.CheckBox1 = True
        Else
            NoteTV.CheckBox1 = False
        End If
        NoteTV.TextBox1 = .Offset(, 1).Value
    End With
End Sub

'''Functions on the Ribbon'''
Sub Out(control As IRibbonControl)
    On Error GoTo Oboy
    ActiveSheet.Activate
    
    NoteTV.Show
Oboy:
End Sub

Sub CoPa(control As IRibbonControl)
Dim Bo1 As Workbook, Bo2 As Workbook
Dim Pa As Worksheet
Set Bo1 = ActiveWorkbook
    On Error GoTo Oboy
Set Pa = Bo1.ActiveSheet
    
    With Bo1
        If .ProtectStructure = True Then
            .Unprotect Environ("userprofile")
        End If
    End With
    With Pa
        If .ProtectContents = True Then
            .Unprotect Environ("userprofile")
        End If
        Set Bo2 = Workbooks.Add
        .Copy Bo2.Worksheets(1)
        .Protect Environ("userprofile")
    End With
    With Bo2
        .ActiveSheet.Protect
    End With
Oboy:
Set Pa = Nothing
Set Bo1 = Nothing
Set Bo2 = Nothing
End Sub

Sub PdfS(control As IRibbonControl)
Dim SWb As Workbook
Dim SWs As Worksheet
Dim FP As String, FNa As String
Set SWb = ActiveWorkbook
    On Error GoTo Oboy
    ActiveSheet.Activate
    On Error GoTo 0
    
Set SWs = SWb.ActiveSheet
    FP = CurDir() & "\PDF_Files"
    If Len(Dir(FP, vbDirectory)) = 0 Then
        MkDir FP
    End If
    FNa = FP & "\" & SWs.Name & "_" & _
    Format(Now, "ddMMyy_hhmmss") & _
    ".pdf"
    With SWs
        .ExportAsFixedFormat xlTypePDF, FNa _
        , xlQualityStandard, OpenAfterPublish:=True
    End With
Oboy:
Set SWb = Nothing
Set SWs = Nothing
FP = vbNullString
FNa = vbNullString
End Sub

Sub Pswd(control As IRibbonControl)
    On Error GoTo Oboy
    ActiveSheet.Activate
    
    MsgBox "WARNING!!!" & Chr(10) & _
    "You are about to set Password" _
    & " for your Active Workbook" _
    & ". Please do not forget it." _
    , vbInformation, "Password Setup"
    
    CPass.Show
Oboy:
End Sub


Function WiAd() As Double
Dim nu1 As Integer, nu2 As Integer
With ActiveCell
    .ColumnWidth = 1
    nu1 = .Width
    .ColumnWidth = 2
    nu2 = .Width - nu1
    WiAd = ((.Height - nu1) / nu2) + 1
End With
End Function

'''UDF (User Define Functions)'''
Function RaCo(Optional RaColumn As String) As Long
    If RaColumn = "" Then
        RaColumn = Replace(ActiveCell.Address, "$", "")
    End If
    On Error Resume Next
    RaCo = Range(RaColumn).Column
    If Err.Number <> 0 Then
        On Error GoTo 0
    End If
    RaColumn = vbNullString
End Function

Function RaRo(Optional RaRow As String) As Long
    If RaRow = "" Then
        RaRow = Replace(ActiveCell.Address, "$", "")
    End If
    On Error Resume Next
    RaRo = Range(RaRow).Row
    If Err.Number <> 0 Then
        On Error GoTo 0
    End If
    RaRow = vbNullString
End Function

Function CelNam(Optional Crow As Long, Optional Ccol As Long) As String
    On Error Resume Next
    If Crow = 0 And Ccol = 0 Then
        CelNam = Replace(ActiveCell.Address, "$", "")
    Else
        CelNam = Replace(Cells(Crow, Ccol).Address, "$", "")
    End If
Crow = 0
Ccol = 0
End Function

Function ActNam(Optional ActSheet As Worksheet) As String
    If ActSheet Is Nothing Then
        Set ActSheet = ActiveSheet
        ActNam = ActSheet.Name
    Else
        ActNam = ActSheet.Name
    End If
Set ActSheet = Nothing
End Function

Function ToRo(Optional tR As Range) As LongPtr
    If tR Is Nothing Then
        Set tR = ActiveCell
        With tR
            ToRo = Cells(Rows.Count, .Column).End(xlUp).Row
        End With
    Else
        With tR
            ToRo = Cells(Rows.Count, .Column).End(xlUp).Row
        End With
    End If
Set tR = Nothing
End Function

Function ToCo(Optional TC As Range) As LongPtr
    If TC Is Nothing Then
        Set TC = ActiveCell
        With TC
            ToCo = Cells(.Row, Columns.Count).End(xlToLeft).Column
        End With
    Else
        With TC
            ToCo = Cells(.Row, Columns.Count).End(xlToLeft).Column
        End With
    End If
Set TC = Nothing
End Function

Function CACel() As Range
Dim Ca As LongPtr
Dim Cbl As LongPtr
Ca = WorksheetFunction.CountA(Cells(1, 1).Resize(ToRo))
Cbl = WorksheetFunction.CountA(Columns(1).Rows)
With ActiveCell
    Set CACel = Cells(ToRo + ((Cbl - Ca) + 1), .Column)
End With
Ca = 0
Cbl = 0
End Function

Function CoS() As Long
Dim CS As Shape
Dim NSh As Long
Dim ShS As Long
Dim ISh As Long
Dim st As Long
Dim k As Long
    st = ActiveCell.Column
    NSh = ActiveSheet.Shapes.Count
    ShS = ActiveCell.Row
    If ActiveCell.Column = 1 Then
    For ISh = ShS To NSh
        Set CS = ActiveSheet.Shapes(ISh)
        With CS
            If Range(.Name).Column > st Then
                k = k + 1
            Else
                Exit For
            End If
        End With
    Next ISh
    Else
    CoS = 0
    End If
    CoS = k
Set CS = Nothing
NSh = 0
ShS = 0
ISh = 0
st = 0
k = 0
End Function

Function ColS() As Integer
Dim CoSt As Integer
Dim VC As Integer
Dim k As Integer
CoSt = ActiveCell.SpecialCells(xlCellTypeLastCell).Column
    With ActiveCell
        If .Column = 1 Then
            For k = .Column To CoSt
                VC = WorksheetFunction. _
                CountA(.Offset(, k) _
                .Resize(CoS))
                If VC = 0 Then
                    Exit For
                End If
            Next k
        End If
    End With
    ColS = k
CoSt = 0
VC = 0
k = 0
End Function

Function fLR() As LongPtr
Dim tR As LongPtr
Dim j As LongPtr
Dim k As LongPtr
    tR = Int(Cells.SpecialCells(xlCellTypeConstants).Count / 2)
    For j = 1 To tR
        If Range("A" & j).EntireRow.Hidden = False Then
            k = 0
            k = Range("A" & j).Row
        End If
    Next j
    fLR = k
tR = 0
j = 0
k = 0
End Function

