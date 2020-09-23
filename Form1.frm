VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.Server svr 
      Left            =   3840
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton CommandBtn 
      Caption         =   "Command1"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.ListBox Variables 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Subs 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim EndSub(0 To 9000) As Boolean
Dim GotoS(0 To 9000) As String
Private Type Variable
Name As String
Data As String
End Type
Dim Vars(0 To 9000) As Variable

Public Function FindSub(SubName As String) As Integer
For I = 1 To Subs.UBound
DoEvents
Sleep 1
 If LCase(VBA.Left(Subs(I).Tag, Len(SubName))) = LCase(SubName) Then FindSub = I: Exit For
Next
End Function

Public Function ReplaceVars(Str As String, Optional SecondReplace As Boolean)
Dim MaxLen As Long
For I = CInt(VarPos("")) - 1 To 0 Step -1
Dim cur As String
cur = Vars(I).Name
If Len(cur) > MaxLen Then MaxLen = Len(cur)
Next
Dim A As Integer
For A = MaxLen To 1 Step -1
 For I = VarPos("") - 1 To 0 Step -1
cur = Vars(I).Name
If Len(cur) = A Then
 If SecondReplace = False Then
 Str = rep(Str, "$" & cur, ReplaceVars(Vars(I).Data, True))
 Else
 Str = rep(Str, "$" & cur, Vars(I).Data)
 End If
End If
Next
Next
ReplaceVars = Str
End Function

Public Function GetVar(Name As String)
GetVar = Vars(VarPos(Name)).Data
End Function

Public Function VarPos(VarName As String) As Integer
For I = 0 To 9000
If LCase(Vars(I).Name) = LCase(VarName) Then VarPos = I: Exit Function
Next
End Function

Public Function SetVar(Name As String, AsS)
Dim AS2 As String
AS2 = CStr(AsS)
AS2 = AS2
If VBA.Left(LCase(AS2), 4) = "left" Or VBA.Left(LCase(AS2), 5) = "right" Or VBA.Left(LCase(AS2), 3) = "mid" Then
AS2 = CheckBrackets(AS2, 0)
End If
AS2 = Process(AS2)
Dim dx As Boolean
Name = ReplaceVars(Name)
Dim CurVar As Integer
CurVar = VarPos(Name)
If LCase(Vars(CurVar).Name) <> LCase(Name) Then CurVar = VarPos("")
With Vars(CurVar)
.Name = Name
.Data = AS2
End With
End Function

Public Function Process(Str As String) As String
On Error Resume Next
If VBA.Left(Str, 3) = "lc(" Then GoTo lcbit
Dim BKSTR As String
BKSTR = CheckBrackets(Str, 0)
If BKSTR <> "" Then
Process = BKSTR
Exit Function
End If
Dim STRTMP As String
If LCase(VBA.Left(Str, 3)) = "tr(" Then
Str = VBA.Mid(Str, 4)
Str = VBA.Left(Str, Len(Str) - 1)
Str = ReplaceVars(Str)
Process = Trim$(Str)
Exit Function
End If
If LCase(VBA.Left(Str, 3)) = "re(" Then
Str = VBA.Mid(Str, 4)
Str = VBA.Left(Str, Len(Str) - 1)
Str = ReplaceVars(Str)
End If
If LCase(VBA.Left(Str, 3)) = "lc(" Then
lcbit:
Str = VBA.Mid(Str, 4)
Str = VBA.Left(Str, Len(Str) - 1)
Str = LCase(Str)
ElseIf LCase(VBA.Left(Str, 4)) = "chr(" Then
Str = VBA.Mid(Str, 5)
Str = VBA.Left(Str, Len(Str) - 1)
Str = Chr(CLng(Val(Str)))
ElseIf LCase(VBA.Left(Str, 3)) = "uc(" Then
Str = VBA.Mid(Str, 4)
Str = VBA.Left(Str, Len(Str) - 1)
Str = UCase(Str)
End If
Process = Str
End Function

Public Function rep(strExpression As String, strFind As String, strReplace As String)
    Dim intX As Integer


    If (Len(strExpression) - Len(strFind)) >= 0 Then


        For intX = 1 To Len(strExpression)


            If LCase(Mid(strExpression, intX, Len(strFind))) = LCase(strFind) Then
                strExpression = Left(strExpression, (intX - 1)) + strReplace + Mid(strExpression, intX + Len(strFind), Len(strExpression))
            End If
        Next
    End If
    rep = strExpression
End Function

Private Sub Command1_Click()
List2.Clear
For I = 0 To VarPos("") - 1
List2.AddItem Vars(I).Name & "=" & Vars(I).Data
Next
End Sub

Private Sub CommandBtn_Click(Index As Integer)
DoSub FindSub("Event_" & CommandBtn(Index).Tag & "_Click")
End Sub

Private Sub CommandBtn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
DoSub FindSub("Event_" & CommandBtn(Index).Tag & "_KeyDown")
With CommandBtn(Index)
SetVar "KeyCode", KeyCode
SetVar "Shift", Shift
End With
End Sub

Private Sub CommandBtn_KeyPress(Index As Integer, KeyAscii As Integer)
DoSub FindSub("Event_" & CommandBtn(Index).Tag & "_KeyPress")
With CommandBtn(Index)
SetVar "KeyCode", KeyAscii
SetVar "Shift", Shift
End With
End Sub

Private Sub CommandBtn_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
DoSub FindSub("Event_" & CommandBtn(Index).Tag & "_KeyUp")
With CommandBtn(Index)
SetVar "KeyCode", KeyCode
SetVar "Shift", Shift
End With
End Sub

Private Sub CommandBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
DoSub FindSub("Event_" & CommandBtn(Index).Tag & "_MouseDown")
With CommandBtn(Index)
SetVar "Button", Button
SetVar "Shift", Shift
SetVar "X", (x / 15)
SetVar "Y", (Y / 15)
End With
End Sub

Private Sub CommandBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
DoSub FindSub("Event_" & CommandBtn(Index).Tag & "_MouseUp")
With CommandBtn(Index)
SetVar "Button", Button
SetVar "Shift", Shift
SetVar "X", (x / 15)
SetVar "Y", (Y / 15)
End With
End Sub

Private Sub Form_Load()
SetVar "Newline", vbNewLine
SetVar "App.Path", App.Path
Open App.Path & "\script.txt" For Input As #1
Dim TMP As String
Dim CurSub As String, CurList As Integer
Dim I As Integer
While Not EOF(1)
Line Input #1, TMP
If Trim$(TMP) = "" Or VBA.Left(TMP, 1) = "#" Then GoTo nextr
 If VBA.Left(LCase(TMP), 4) = "sub " Then
 I = 0
 Load Subs(Subs.UBound + 1)
 CurList = Subs.UBound
 CurSub = VBA.Mid(TMP, 5)
 Subs(CurList).Tag = CurSub
 ElseIf LCase(TMP) = "end sub" Then
 I = 0
 CurSub = ""
 CurList = 0
 Else
 I = I + 1
 If VBA.Right(TMP, 1) = ":" Then GoTo OkieDokey
 If Not VBA.Right(TMP, 1) = ";" And InStr(TMP, ";") = 0 Then
 DebugList "Syntax Error: Missing "";"" on Line " & I & " of " & Subs(CurList).Tag & "."
 ElseIf VBA.Left(TMP, 1) = ";" Then
 Else
 TMP = VBA.Left(TMP, InStrRev(TMP, ";") - 1)
OkieDokey:
  If VBA.Right(TMP, 1) = ":" Then
   If LCase(TMP) Like "*.*:" Then
    DebugList "Syntax Error: Dot in label ""."" on Line " & I & " of " & Subs(CurList).Tag & "."
   Else
   Subs(CurList).AddItem TMP
   End If
  Else
  Subs(CurList).AddItem TMP
  End If
 End If
 End If
nextr:
Wend
DoSub FindSub("main")
Close #1
End Sub

Public Function DebugList(Str As String)
List1.AddItem Str
List1.ListIndex = List1.ListCount - 1
End Function

Public Function DoSub(SubIndex As Integer)
Dim A As Integer
For I = 0 To Subs(SubIndex).ListCount - 1

If EndSub(SubIndex) = True Then
EndSub(SubIndex) = False
Exit Function
End If

If LCase(Trim$(GotoS(SubIndex))) <> "" Then
GotoS(SubIndex) = Trim$(GotoS(SubIndex))
For A = 0 To Subs(SubIndex).ListCount - 1
If VBA.Right(Subs(SubIndex).List(A), 1) = ":" Then
 If LCase(VBA.Left(Subs(SubIndex).List(A), Len(Subs(SubIndex).List(A)) - 1)) = LCase(GotoS(SubIndex)) Then
 GotoS(SubIndex) = ""
 I = (A - 1)
 GoTo nextpoint
 End If
End If
Next
End If

CheckLine Subs(SubIndex).List(I), CInt(SubIndex)
nextpoint:
Next
End Function

Public Function CheckLine(Str As String, ListIndex As Integer)
 Dim Var2Set As String
 Dim VarValue As String
On Error Resume Next
 SubCommand Str, ListIndex
 If LCase(VBA.Left(Str, InStr(Str, " ") - 1)) = "goto" Then SubCommand Str, ListIndex
 If VBA.Mid(Str, InStr(Str, " ") + 1, 1) = "=" And VBA.Left(Str, 1) = "$" Then
  Var2Set = VBA.Left(Str, InStr(Str, " ") - 1)
  Var2Set = VBA.Mid(Var2Set, 2)
  VarValue = VBA.Mid(Str, InStr(Str, " ") + 1)
  VarValue = VBA.Mid(VarValue, InStr(VarValue, " ") + 1)
  VarValue = CheckQuotes(VarValue)
  VarValue = ReplaceVars(VarValue)
  SetVar CStr(Var2Set), CStr(VarValue)
 Else
 CheckBrackets Str, ListIndex
 End If
End Function

Public Function SubCommand(Str As String, ListIndex As Integer)
On Error GoTo err
CheckBrackets Str, ListIndex
Select Case LCase(VBA.Left(Str, InStr(Str, " ") - 1))
Case "dosub"
DoSub FindSub(CheckQuotes(VBA.Mid(Str, InStr(Str, " ") + 1), True))
Case "exit"
If LCase(VBA.Mid(Str, InStr(Str, " ") + 1)) = "sub" Then
EndSub(ListIndex) = True
Else
EndSub(FindSub(CheckQuotes(VBA.Mid(Str, InStr(Str, " ") + 1), True))) = True
End If
End Select
err:
End Function

Public Function LoadFile(Str As String) As String '
On Error GoTo err
Open Str For Binary As #3
Str = Input(LOF(3), 3)
LoadFile = Str
err:
Close #3
End Function

Public Function CheckBrackets(Str As String, ListIndex As Integer)
Dim DaTemp As String, Result As String, VarValue As String, Var2Set As String, VarA As String, VarB As String
Dim SpltTmp As Variant, Blanks As Integer, WasDot As Boolean

Dim InputVars_0 As String, InputVars_1 As String, InputVars_2 As String

Blanks = 0
'On Error Resume Next
Dim Str2 As String
Str2 = VBA.Left(Str, 14)
If Not Str2 Like "*(*" Then
Exit Function
End If
Select Case VBA.Left(LCase(Str), InStr(Str, "(") - 1)
Case "lc"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
CheckBrackets = LCase(Str)
Exit Function
Case "uc"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
CheckBrackets = UCase(Str)
Exit Function
Case "tr"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
CheckBrackets = Trim$(Str)
Exit Function
Case "chr"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
CheckBrackets = Chr(Str)
Exit Function
Case "asc"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
CheckBrackets = Asc(Str)
Exit Function
Case "replace"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
VarValue = VBA.Left(Str, InStr(Str, ",") - 1)
Str = VBA.Mid(Str, InStr(Str, ",") + 1)
VarA = VBA.Left(Str, InStr(Str, ",") - 1)
Var2Set = VBA.Mid(Str, InStr(Str, ",") + 1)
VarValue = CheckQuotes(VarValue, True)
Var2Set = CheckQuotes(Var2Set, True)
VarA = CheckQuotes(VarA, True)
CheckBrackets = rep(VarValue, VarA, Var2Set)
Exit Function
Case "loadfile"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
Str = LoadFile(Str)
If Str = "" Then Str = "File Not Found"
CheckBrackets = Str
Exit Function
Case "closesocket"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
svr.CloseSocket CInt(Str)
Case "senddata"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
VarValue = CheckQuotes(VBA.Left(Str, InStr(Str, ",") - 1), True)
Var2Set = CheckQuotes(VBA.Mid(Str, InStr(Str, ",") + 1), True)
svr.SendData Var2Set, CLng(VarValue)
Case "instr"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
VarValue = CheckQuotes(VBA.Left(Str, InStr(Str, ",") - 1), True)
Var2Set = CheckQuotes(VBA.Mid(Str, InStr(Str, ",") + 1), True)
CheckBrackets = InStr(Var2Set, VarValue)
Exit Function
Case "instrrev"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
VarValue = CheckQuotes(VBA.Left(Str, InStr(Str, ",") - 1), True)
Var2Set = CheckQuotes(VBA.Mid(Str, InStr(Str, ",") + 1), True)
CheckBrackets = InStrRev(Var2Set, VarValue)
Exit Function
Case "left"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
VarValue = CheckQuotes(VBA.Left(Str, InStr(Str, ",") - 1), True)
Var2Set = CheckQuotes(VBA.Mid(Str, InStr(Str, ",") + 1), True)
CheckBrackets = VBA.Left(Var2Set, CLng(VarValue))
Exit Function
Case "mid"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
VarValue = CheckQuotes(VBA.Left(Str, InStr(Str, ",") - 1), True)
Var2Set = CheckQuotes(VBA.Mid(Str, InStr(Str, ",") + 1), True)
CheckBrackets = VBA.Mid(Var2Set, CLng(VarValue))
Exit Function
Case "right"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
VarValue = CheckQuotes(VBA.Left(Str, InStr(Str, ",") - 1), True)
Var2Set = CheckQuotes(VBA.Mid(Str, InStr(Str, ",") + 1), True)
CheckBrackets = VBA.Right(Var2Set, CLng(VarValue))
Exit Function
Case "startserver"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
svr.StartServer CLng(Str)
Case "stopserver"
svr.StopServer
Case "goto"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
If Str Like "*.*" Then
 ListIndex = FindSub(VBA.Left(Str, InStr(Str, ".") - 1))
 Str = VBA.Mid(Str, InStr(Str, ".") + 1)
 WasDot = True
End If
GotoS(ListIndex) = Str
If WasDot = True Then DoSub ListIndex
Case "msgbox"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
MsgBox Str
Case "inputbox"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
SpltTmp = SplitLine(Str, """,""")
InputVars_0 = SpltTmp(0) & """"
InputVars_1 = """" & VBA.Mid(SpltTmp(1), 3) & """"
InputVars_2 = VBA.Mid(SpltTmp(2), 2)
InputVars_0 = CheckQuotes(InputVars_0)
InputVars_1 = CheckQuotes(InputVars_1)
InputVars_2 = CheckQuotes(InputVars_2)
CheckBrackets = InputBox(InputVars_0, InputVars_1, InputVars_2)
Exit Function
Case "printdebug"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str, True)
DebugList Str
Case "createbtn"
Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
Str = CheckQuotes(Str)
Str2 = Str

SpltTmp = SplitLine(Str)

For I = 0 To 4
SpltTmp(I) = ReplaceVars(CheckQuotes(CStr(SpltTmp(I)), False))
If LCase(Trim$(CStr(SpltTmp(I)))) = "" Then Blanks = Blanks + 1
Next

If Blanks = 5 Then
 SpltTmp = SplitLine(Str2)
For I = 0 To 4
SpltTmp(I) = ReplaceVars(CStr(SpltTmp(I)))
Next
End If

If CheckName(CStr(GetVar("btnname"))) = True Or Trim$(LCase(CStr(GetVar("btnname")))) = "" Then
DebugList "Name: " & CStr(GetVar("btnname")) & " already exists!"
Exit Function
End If

Load CommandBtn(CommandBtn.UBound + 1)
With CommandBtn(CommandBtn.UBound)
.Left = (SpltTmp(0) * 15)
.Top = (SpltTmp(1) * 15)
.Width = (SpltTmp(2) * 15)
.Height = (SpltTmp(3) * 15)
.Tag = CStr(GetVar("btnname"))
.Caption = SpltTmp(4)
.Visible = True
.TabStop = True
End With
Case "if"
 DaTemp = VBA.Mid(Str, InStr(Str, "(") + 1)
 DaTemp = VBA.Left(DaTemp, Len(DaTemp) - 1)
 Var2Set = VBA.Mid(DaTemp, InStr(DaTemp, " ") + 1)
 VarValue = VBA.Mid(Var2Set, InStr(Var2Set, " ") + 1)
 Result = VBA.Mid(VarValue, InStr(VarValue, " ") + 1)
 Result = VBA.Mid(Result, InStr(Result, " ") + 1)
 VarValue = VBA.Left(VarValue, InStr(VarValue, " ") - 1)
 DaTemp = VBA.Left(DaTemp, InStr(DaTemp, " then") - 1)
 Var2Set = VBA.Mid(DaTemp, InStr(DaTemp, " ") + 1)
 VarValue = VBA.Mid(Var2Set, InStr(Var2Set, " ") + 1)
 Var2Set = VBA.Left(Var2Set, InStr(Var2Set, " ") - 1)
 DaTemp = VBA.Left(DaTemp, InStr(DaTemp, " ") - 1)
 DaTemp = ReplaceVars(DaTemp)
 VarValue = ReplaceVars(VarValue)
 DaTemp = Process(DaTemp)
 VarValue = Process(VarValue)
 DaTemp = CheckQuotes(DaTemp)
 VarValue = CheckQuotes(VarValue)
 
  
  Select Case LCase(Var2Set)
  Case "="
   If DaTemp = VarValue Then SubCommand Result, ListIndex
  Case "<>"
   If DaTemp <> VarValue Then SubCommand Result, ListIndex
  Case "<"
   If Val(DaTemp) < Val(VarValue) Then SubCommand Result, ListIndex
  Case ">"
   If Val(DaTemp) > Val(VarValue) Then SubCommand Result, ListIndex
  Case ">="
   If Val(DaTemp) >= Val(VarValue) Then SubCommand Result, ListIndex
  Case "<="
   If Val(DaTemp) <= Val(VarValue) Then SubCommand Result, ListIndex
  Case "like"
   If DaTemp Like VarValue Then
   SubCommand Result, ListIndex
   End If
  End Select
  
'NUMBER FUNCTIONS

 Case "add"
 Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
 VarValue = Str
 VarA = VBA.Left(VarValue, InStr(VarValue, ",") - 1)
 VarB = VBA.Mid(VarValue, InStr(VarValue, ",") + 1)
 VarName = VarA
 VarA = Process(CStr(VarA))
 VarB = Process(CStr(VarB))
 VarA = CheckQuotes(VarA, True)
 VarB = CheckQuotes(VarB, True)
 VarA = Val(VarA) + Val(VarB)
 SetVar VBA.Mid(VarName, 2), CStr(VarA)
 Case "mod"
 Str = VBA.Mid(Str, InStr(Str, "(") + 1)
 Str = VBA.Left(Str, Len(Str) - 1)
 VarValue = Str
 VarA = VBA.Left(VarValue, InStr(VarValue, ",") - 1)
 VarB = VBA.Mid(VarValue, InStr(VarValue, ",") + 1)
 VarName = VarA
 VarA = Process(CStr(VarA))
 VarB = Process(CStr(VarB))
 VarA = CheckQuotes(VarA, True)
 VarB = CheckQuotes(VarB, True)
 VarA = Val(VarA) Mod Val(VarB)
 SetVar VBA.Mid(VarName, 2), CStr(VarA)
 Case "multiply"
 Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
 VarValue = Str
 VarA = VBA.Left(VarValue, InStr(VarValue, ",") - 1)
 VarB = VBA.Mid(VarValue, InStr(VarValue, ",") + 1)
 VarName = VarA
 VarA = Process(CStr(VarA))
 VarB = Process(CStr(VarB))
 VarA = CheckQuotes(VarA, True)
 VarB = CheckQuotes(VarB, True)
 VarA = Val(VarA) * Val(VarB)
 SetVar VBA.Mid(VarName, 2), CStr(VarA)
 Case "divide"
 Str = VBA.Mid(Str, InStr(Str, "(") + 1)
Str = VBA.Left(Str, Len(Str) - 1)
 VarValue = Str
 VarA = VBA.Left(VarValue, InStr(VarValue, ",") - 1)
 VarB = VBA.Mid(VarValue, InStr(VarValue, ",") + 1)
 VarName = VarA
 VarA = Process(CStr(VarA))
 VarB = Process(CStr(VarB))
 VarA = CheckQuotes(VarA, True)
 VarB = CheckQuotes(VarB, True)
 VarA = Val(VarA) + Val(VarB)
 SetVar VBA.Mid(VarName, 2), CStr(VarA)
 Case "minus"
 Str = VBA.Mid(Str, InStr(Str, "(") + 1)
 Str = VBA.Left(Str, Len(Str) - 1)
 VarValue = Str
 VarA = VBA.Left(VarValue, InStr(VarValue, ",") - 1)
 VarB = VBA.Mid(VarValue, InStr(VarValue, ",") + 1)
 VarName = VarA
 VarA = Process(CStr(VarA))
 VarB = Process(CStr(VarB))
 VarA = CheckQuotes(VarA, True)
 VarB = CheckQuotes(VarB, True)
 VarA = Val(VarA) - Val(VarB)
 SetVar VBA.Mid(VarName, 2), CStr(VarA)
End Select
err:
CheckBrackets = Str
End Function

Public Function CheckName(Str As String) As Boolean
For I = 1 To CommandBtn.UBound
If LCase(CommandBtn(I).Tag) = LCase(Str) Then CheckName = True: Exit Function
Next
End Function

Public Function CheckQuotes(Str As String, Optional ReplaceTheVars As Boolean) As String
On Error GoTo err
If VBA.Left(Str, 1) = Chr(34) Then
Str = VBA.Mid(Str, 2)
Str = VBA.Left(Str, Len(Str) - 1)
Str = Replace(Str, Chr(34) & Chr(34), Chr(34))
If ReplaceTheVars = True Then Str = ReplaceVars(Str)
CheckQuotes = Str
Else
If ReplaceTheVars = True Then Str = ReplaceVars(Str)
CheckQuotes = Str
End If
err:
End Function

Public Function SplitLine(TMP As String, Optional Splitter As String)
If Splitter = "" Then Splitter = ","
Dim Instance As Long
For I = 1 To Len(TMP) Step 1
If LCase(Mid(TMP, I, Len(Splitter))) = LCase(Splitter) Then Instance = Instance + 1
Next
If Instance = 0 Then SplitLine = TMP: Exit Function
Dim tmpx(1 To 100) As String
Dim li As Integer
For I = 1 To Instance
tmpx(I) = VBA.Left(TMP, InStr(TMP, Splitter) - 1)
tmpx(I) = tmpx(I)
TMP = Mid(TMP, InStr(TMP, Splitter) + 1)
li = I
Next
li = li + 1
tmpx(li) = TMP
SplitLine = Array(tmpx(1), tmpx(2), tmpx(3), tmpx(4), tmpx(5), tmpx(6), tmpx(7), tmpx(8), tmpx(9), tmpx(10), tmpx(11), tmpx(12), tmpx(13), tmpx(14), tmpx(15), tmpx(16), tmpx(17), tmpx(18), tmpx(19), tmpx(20), tmpx(21), tmpx(22), tmpx(23), tmpx(24), tmpx(25), tmpx(26), tmpx(27), tmpx(28), tmpx(29), tmpx(30), tmpx(31), tmpx(32), tmpx(33), tmpx(34), tmpx(35), tmpx(36), tmpx(37), tmpx(38), tmpx(39), tmpx(40), tmpx(41), tmpx(42), tmpx(43), tmpx(44), tmpx(45), tmpx(46), tmpx(47), tmpx(48), tmpx(49), tmpx(50), tmpx(51), tmpx(52), tmpx(53), tmpx(54), tmpx(55), tmpx(56), tmpx(57), tmpx(58), tmpx(59), tmpx(60), tmpx(61), tmpx(62), tmpx(63), tmpx(64), tmpx(65), tmpx(66), tmpx(67), tmpx(68), tmpx(69), tmpx(70), tmpx(71), tmpx(72), tmpx(73), tmpx(74), tmpx(75), tmpx(76), tmpx(77), tmpx(78), tmpx(79), tmpx(80), tmpx(81), tmpx(82), tmpx(83), tmpx(84), tmpx(85), tmpx(86), tmpx(87), tmpx(88), tmpx(89), tmpx(90), tmpx(91), tmpx(92), tmpx(93), tmpx(94), tmpx(95), tmpx(96), tmpx(97), tmpx(98), tmpx(99), tmpx(100))
End Function

Private Sub Form_Unload(Cancel As Integer)
DoSub FindSub("StopServer")
End Sub

Private Sub List2_DblClick()
MsgBox List2.Text
End Sub

Private Sub svr_DataArrival(ByVal idx As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
SetVar "Server_Idx", idx
SetVar "Server_Data", Data
SetVar "Server_BytesTotal", bytesTotal
SetVar "Server_RemoteIP", RemoteIP
SetVar "Server_RemoteHost", RemoteHost
DoSub FindSub("Server_DataArrival")
End Sub

Private Sub svr_Error(ByVal idx As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
SetVar "Server_StartInfo", "Could not start on port #" & svr.ServerPort
DoSub FindSub("Server_StartInfo")
End Sub

Private Sub svr_SendComplete(idx As Integer)
SetVar "Server_Idx", idx
DoSub FindSub("Server_SendComplete")
End Sub

Private Sub svr_ServerStarted()
SetVar "Server_StartInfo", "Started on port #" & svr.ServerPort & "."
DoSub FindSub("Server_StartInfo")
End Sub

Private Sub svr_ServerStopped()
SetVar "Server_StartInfo", "Stopped."
DoSub FindSub("Server_StartInfo")
End Sub
