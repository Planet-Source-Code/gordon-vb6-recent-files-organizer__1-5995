VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3435
   ClientLeft      =   3045
   ClientTop       =   1740
   ClientWidth     =   7440
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2190
      Width           =   6770
      Begin VB.CommandButton cmdDown 
         Caption         =   "¯"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdUP 
         Caption         =   "­"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   6
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "?"
         Height          =   315
         Left            =   6150
         TabIndex        =   5
         Top             =   270
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         MaskColor       =   &H000000FF&
         TabIndex        =   2
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4230
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelall 
         Caption         =   "Delete &all and exit program"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         Top             =   270
         Width           =   2115
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   1815
      Index           =   0
      Left            =   240
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   360
      Width           =   6770
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   1815
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   6770
   End
   Begin VB.Label Label2 
      Caption         =   "By: Gordon Li"
      Height          =   225
      Left            =   750
      TabIndex        =   11
      Top             =   3180
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Organizer"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   10
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyReg As New cReadWriteEasyReg
Dim Change As Boolean
Dim intIndex As Integer
Dim MyVariant As Variant

Private Sub cmdAbout_Click()
frmAbout.Show 1
End Sub

Private Sub cmdApply_Click()
If List1(intIndex).ListCount > 0 And Change Then
For i = LBound(MyVariant) To UBound(MyVariant)
    MyReg.DeleteValue (MyVariant(Str(i)))
Next
For i = 0 To List1(intIndex).ListCount - 1
    Call MyReg.CreateValue(i + 1, List1(intIndex).List(i), REG_SZ)
Next
End If
cmdApply.Enabled = False
'MyReg.CloseRegistry
End Sub
Private Sub cmdDelall_Click()
Change = True
removeList (0)
For i = LBound(MyVariant) To UBound(MyVariant)
    MyReg.DeleteValue (MyVariant(Str(i)))
Next
Unload Me
End Sub

Private Sub cmdDelete_Click()
With List1(intIndex)
If .ListIndex = -1 Then Exit Sub
If .SelCount = 1 Then
.RemoveItem .ListIndex
Else
    For i = .ListCount - 1 To 0 Step -1
        If .Selected(i) = True Then
            .RemoveItem i
        End If
    Next
End If

Change = True
cmdApply.Enabled = True
cmdDelete.Enabled = False
If .ListCount = 0 Then
    cmdDelall.Enabled = False
End If
End With
End Sub

Private Sub Form_Load()
Dim regKey As String
Form1.Caption = App.Comments
cmdUP.Visible = False
cmdDown.Visible = False
regKey = ".DEFAULT\Software\Microsoft\Visual Basic\6.0\RecentFiles"
'load the content from registry to Listbox
If Not MyReg.OpenRegistry(HKEY_USERS, regKey) Then
    MsgBox "Couldn't open the registry"
    End
    Exit Sub
End If
MyVariant = MyReg.GetAllValues
On Error Resume Next
For i = LBound(MyVariant) To UBound(MyVariant)
    List1(0).AddItem MyReg.GetValue(MyVariant(Str(i)))
    List1(1).AddItem List1(0).List(i)
Next i
Change = False
If List1(0).ListCount = 0 Then
    res = MsgBox("There's nothing in the recent list" & Chr(13) & "Press ok to exit", vbOKOnly, "Open error")
    If res = 1 Then
        Unload Me
        Exit Sub
    End If
Else
    List1(0).ListIndex = 0
    cmdDelall.Enabled = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MyReg.CloseRegistry
End Sub

Private Sub Label1_Click(Index As Integer)
intIndex = Index

Label1(Index).FontBold = True
If Change Then
    removeList (Index)
    addlist (Index)
End If

If List1(Index).SelCount <> 0 Then
    cmdDelete.Enabled = True
Else
    cmdDelete.Enabled = False
End If
Select Case Index
Case 0
    Label1(1).FontBold = False
    List1(0).Visible = True
    List1(1).Visible = False
    cmdUP.Visible = False
    cmdDown.Visible = False
Case 1
    'List1(0).ListIndex = -1
    Label1(0).FontBold = False
    List1(0).Visible = False
    List1(1).Visible = True
    cmdUP.Visible = True
    cmdDown.Visible = True
End Select
End Sub

Private Sub removeList(Index As Integer)
For i = List1(Index).ListCount - 1 To 0 Step -1
    List1(Index).RemoveItem i
Next
End Sub

Private Sub addlist(Index As Integer)
Dim temp As Byte
    If Index = 0 Then
        temp = 1
    ElseIf Index = 1 Then
        temp = 0
    End If
    For i = 0 To List1(temp).ListCount - 1
        List1(Index).AddItem List1(temp).List(i)
    Next
End Sub

Private Sub cmdUP_Click()
UpDown (-1)
End Sub

Private Sub cmdDown_Click()
UpDown 1
End Sub

Private Sub UpDown(x As Integer)
Change = True
cmdApply.Enabled = True
Dim temp As String
With List1(1)
    temp = .List(.ListIndex)
    .List(.ListIndex) = .List(.ListIndex + x)
    .List(.ListIndex + x) = temp
    .Selected(.ListIndex + x) = True
End With
End Sub

Private Sub List1_Click(Index As Integer)
    cmdUP.Enabled = False
    cmdDown.Enabled = False
    With Me.List1(Index)
    If .ListCount > 1 Then
    If .SelCount > 0 Then
        cmdDelete.Default = True
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
        If .ListIndex > 0 And .ListIndex < .ListCount - 1 Then
            cmdUP.Enabled = True
            cmdDown.Enabled = True
        End If
        If .ListIndex = 0 Then cmdDown.Enabled = True
        If .ListIndex = .ListCount - 1 Then cmdUP.Enabled = True
        If .SelCount > 1 Then
            cmdUP.Enabled = False
            cmdDown.Enabled = False
        End If
    End If
    End With
End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
List1_Click (Index)
End Sub


