VERSION 5.00
Begin VB.Form SCForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scentific Calculator Beta 1"
   ClientHeight    =   6975
   ClientLeft      =   6510
   ClientTop       =   2310
   ClientWidth     =   6150
   FillStyle       =   0  'Solid
   Icon            =   "ProjectScientificCalculatorForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6150
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "            x      10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   15
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "           y      X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   14
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "           3      X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   13
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "           2      X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   12
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "      y  __     V x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   11
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "      3  __     V x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   10
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "      2  __     V x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "          1                  ---         x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "       log                  10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ln"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "        log                y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "mod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "mr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "m-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "m+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "mc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   19
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   18
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   17
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandDev 
      BackColor       =   &H8000000D&
      Caption         =   "Developers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CommandBackSpace 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CommandPercent 
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CommandMorP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CommandScientific 
      BackColor       =   &H8000000D&
      Caption         =   "Sientific calc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CommandS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   16
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CommandDecimal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton CommandExit 
      BackColor       =   &H8000000D&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CommandClear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CommandDivide 
      BackColor       =   &H000080FF&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CommandMultiply 
      BackColor       =   &H000080FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CommandSubtract 
      BackColor       =   &H000080FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CommandAdd 
      BackColor       =   &H000080FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton CommandEquals 
      BackColor       =   &H000080FF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   1095
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      Height          =   1095
      Index           =   9
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      Height          =   1095
      Index           =   8
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      Height          =   1095
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      Height          =   1095
      Index           =   6
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      Height          =   1095
      Index           =   5
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      Height          =   1095
      Index           =   2
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      Height          =   1095
      Index           =   3
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      Height          =   1095
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton CommandNum 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      Height          =   1095
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label panel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "SCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n0 As Single
Dim n1 As Single
Dim n2 As Single
Dim op As String
Dim Sop As String
Dim memorizedNum As Double
Dim opPendingAdd As Boolean
Dim opPendingSubtract As Boolean
Dim opPendingMultiply As Boolean
Dim opPendingDivide As Boolean
Dim newNumberEquals As Boolean
Dim newNumber As Boolean

Private Sub CommandNum_Click(Index As Integer)

If panel.Caption = "0" Or newNumber = True Or panel.Caption = "Error" Then

newNumber = False

n1 = Val(panel.Caption)

n2 = 0

opPendingAdd = False
opPendingSubtract = False
opPendingDivide = False
opPendingMultiply = False

panel.Caption = ""

End If

If newNumberEquals = True Then

newNumberEquals = False

panel.Caption = ""

n1 = 0

End If

If Len(panel.Caption) < 20 Then

panel.Caption = panel.Caption + Right$(Str(Index), 1)

If Len(panel.Caption) > 16 Then

panel.Font.Size = panel.Font.Size - 1.5

Else

panel.FontSize = 26

End If

Else

panel.Caption = panel.Caption

End If

End Sub

Private Sub CommandDecimal_Click()

If panel.Caption = "Developed by Ahmad Aussaili and Seren Wassouf" Then
panel.Caption = "0"
panel.FontSize = 26
End If

If InStr(panel.Caption, ".") = 0 Then
panel.Caption = panel.Caption + "."
End If

If newNumber = True Then

newNumber = False

panel.Caption = "0."

End If

End Sub

Private Sub CommandAdd_Click()

newNumberEquals = False

newNumber = True

opPendingAdd = True

If op = "-" Then

If opPendingSubtract = True Then
opPendingSubtract = False
panel.Caption = panel.Caption
op = "+"
Else
op = "+"
n2 = Val(panel.Caption)
panel.Caption = n1 - n2
End If

ElseIf op = "*" Then

If opPendingMultiply = True Then
opPendingMultiply = False
panel.Caption = panel.Caption
op = "+"
Else
op = "+"
n2 = Val(panel.Caption)
panel.Caption = n1 * n2
End If

ElseIf op = "/" Then

If opPendingDivide = True Then
opPendingDivide = False
panel.Caption = panel.Caption
op = "+"
Else
op = "+"
n2 = Val(panel.Caption)

If n2 <> 0 Then
panel.Caption = n1 / n2
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

End If

Else

If n1 <> 0 Then
op = "+"
n2 = Val(panel.Caption)
panel.Caption = n1 + n2
n1 = 0
Else
op = "+"
n1 = Val(panel.Caption)
panel.Caption = panel.Caption
End If

End If

End Sub

Private Sub CommandSubtract_Click()

newNumber = True

newNumberEquals = False

opPendingSubtract = True

If op = "+" Then

If opPendingAdd = True Then
opPendingAdd = False
panel.Caption = panel.Caption
op = "-"
Else
op = "-"
n2 = Val(panel.Caption)
panel.Caption = n1 + n2
End If

ElseIf op = "*" Then

If opPendingMultiply = True Then
opPendingMultiply = False
panel.Caption = panel.Caption
op = "-"
Else
op = "-"
n2 = Val(panel.Caption)
panel.Caption = n1 * n2
End If

ElseIf op = "/" Then

If opPendingDivide = True Then
opPendingDivide = False
panel.Caption = panel.Caption
op = "-"
Else
op = "-"
n2 = Val(panel.Caption)

If n2 <> 0 Then
panel.Caption = n1 / n2
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

End If

Else

If n1 <> 0 Then
op = "-"
n2 = Val(panel.Caption)
panel.Caption = n1 - n2
n1 = 0
Else
op = "-"
n1 = Val(panel.Caption)
panel.Caption = panel.Caption
End If

End If

End Sub

Private Sub CommandMultiply_Click()

newNumber = True

newNumberEquals = False

opPendingMultiply = True

If op = "+" Then

If opPendingAdd = True Then
opPendingAdd = False
panel.Caption = panel.Caption
op = "*"
Else
op = "*"
n2 = Val(panel.Caption)
panel.Caption = n1 + n2
End If

ElseIf op = "-" Then

If opPendingSubtract = True Then
opPendingSubtract = False
panel.Caption = panel.Caption
op = "*"
Else
op = "*"
n2 = Val(panel.Caption)
panel.Caption = n1 - n2
End If

ElseIf op = "/" Then

If opPendingDivide = True Then
opPendingDivide = False
panel.Caption = panel.Caption
op = "*"
Else
op = "*"
n2 = Val(panel.Caption)

If n2 <> 0 Then
panel.Caption = n1 / n2
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

End If

Else

If n1 <> 0 Then
op = "*"
n2 = Val(panel.Caption)
panel.Caption = n1 * n2
n1 = 0
Else
op = "*"
n1 = Val(panel.Caption)
panel.Caption = panel.Caption
End If

' or If n2 = 0 Then

'op = "*"
'n2 = Val(panel.Caption)
'panel.Caption = n1 * n2

'Else

'panel.Caption = panel.Caption

'End If

End If

End Sub

Private Sub CommandDivide_Click()

newNumber = True

newNumberEquals = False

opPendingDivide = True

If op = "+" Then

If opPendingAdd = True Then
opPendingAdd = False
panel.Caption = panel.Caption
op = "/"
Else
op = "/"
n2 = Val(panel.Caption)
panel.Caption = n1 + n2
End If

ElseIf op = "-" Then

If opPendingSubtract = True Then
opPendingSubtract = False
panel.Caption = panel.Caption
op = "/"
Else
op = "/"
n2 = Val(panel.Caption)
panel.Caption = n1 - n2
End If

ElseIf op = "*" Then

If opPendingMultiply = True Then
opPendingMultiply = False
panel.Caption = panel.Caption
op = "/"
Else
op = "/"
n2 = Val(panel.Caption)
panel.Caption = n1 * n2
End If

Else

If n1 <> 0 Then
op = "/"
n2 = Val(panel.Caption)

If n2 <> 0 Then
panel.Caption = n1 / n2
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

n1 = 0
Else
op = "/"
n1 = Val(panel.Caption)
panel.Caption = panel.Caption
End If

End If

End Sub

Private Sub CommandEquals_Click()

newNumberEquals = True

If panel.Caption = "Developed by Ahmad Aussaili and Seren Wassouf" Then
panel.Caption = "0"
panel.FontSize = 26
End If

If Len(panel.Caption) > 16 Then

panel.Font.Size = 20

End If

n2 = Val(panel.Caption)

If op = "+" Then
panel.Caption = n1 + n2
n1 = 0

ElseIf op = "-" Then
panel.Caption = n1 - n2
n1 = 0

ElseIf op = "*" Then
panel.Caption = n1 * n2
n1 = 0

ElseIf op = "/" Then

If n2 <> 0 Then
panel.Caption = n1 / n2
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

n1 = 0

End If
 
If Sop = "x^y" Then

If n0 < 0 And n2 < 1 Then
panel.Caption = "Error"
panel.FontSize = 26
Else
panel.Caption = n0 ^ n2
End If

ElseIf Sop = "x^(1/y)" Then

If n2 = 0 Then
panel.Caption = "Error"
panel.FontSize = 26
End If

'If n2 Mod 2 = 0 Then

If n0 >= 0 Then
panel.Caption = n0 ^ (1 / n2)
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

'Else

'panel.Caption = n0 ^ (1 / n2)

'End If

ElseIf Sop = "logyx" Then

If n0 > 0 And n2 > 0 Then
panel.Caption = Log(n0) / Log(n2)
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

ElseIf Sop = "mod" Then

If n2 <> 0 Then
panel.Caption = n0 Mod n2
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

End If

op = ""
Sop = ""

End Sub

Private Sub CommandClear_Click()

n0 = 0
n1 = 0
n2 = 0

panel.Caption = "0"

panel.FontSize = 26

End Sub

Private Sub CommandMorP_Click()

If panel.Caption = "Developed by Ahmad Aussaili and Seren Wassouf" Then
panel.Caption = "0"
panel.FontSize = 26
End If

If panel.Caption <> 0 Then
 If Left$(panel.Caption, 1) = "-" Then
                    panel.Caption = Right$(panel.Caption, Len(panel.Caption) - 1)
                Else
                    panel.Caption = "-" & panel.Caption
                End If
End If

End Sub

Private Sub CommandPercent_Click()

newNumber = True

n1 = Val(panel.Caption)

panel.Caption = Round(n1 / 100, 20)

If Len(panel.Caption) < 16 Then
panel.Font.Size = 26
Else
panel.FontSize = 20
End If

End Sub

Private Sub CommandBackSpace_Click()

If panel.Caption = "Developed by Ahmad Aussaili and Seren Wassouf" Then
panel.Caption = "0"
panel.FontSize = 26
End If

panel.Caption = Left$(panel.Caption, Len(panel.Caption) - 1)

If panel.Caption = "" Or panel.Caption = "-" Then
panel.Caption = "0"
End If

If Len(panel.Caption) = 16 Then

panel.FontSize = 26

End If

End Sub

Private Sub CommandScientific_Click()

Dim button As CommandButton

For Each button In CommandS

If button.Visible = True Then

button.Visible = False

SCForm.Height = 7395
SCForm.Width = 6240

Else

button.Visible = True

SCForm.Height = 7395
SCForm.Width = 11025

End If

Next

End Sub

Private Sub CommandDev_Click()

panel.Caption = "Developed by Ahmad Aussaili and Seren Wassouf"

panel.FontSize = 12

End Sub

Private Sub CommandExit_Click()
End
End Sub

Private Sub CommandS_Click(Index As Integer)

newNumber = True

If panel.Caption = "Developed by Ahmad Aussaili and Seren Wassouf" Then
panel.Caption = "0"
panel.FontSize = 26
End If

If Len(panel.Caption) > 14 Then
panel.FontSize = 20
Else
panel.FontSize = 26
End If

If Index <> 3 Then       'mr'
n0 = Val(panel.Caption)
End If

If Index = 0 Then                                'mc'

memorizedNum = 0
 
ElseIf Index = 1 Then                            'm+'

memorizedNum = memorizedNum + n0
 
ElseIf Index = 2 Then                            'm-'

memorizedNum = memorizedNum - n0

ElseIf Index = 3 Then                            'mr'

If Len(memorizedNum) < 14 Then

panel.FontSize = 26

Else

panel.FontSize = 20

End If

panel.Caption = memorizedNum

ElseIf Index = 4 Then                            'mod'

Sop = "mod"

panel.Caption = "0"

panel.FontSize = 26

ElseIf Index = 5 Then                            'logxy'

Sop = "logyx"

panel.Caption = "0"

panel.FontSize = 26

ElseIf Index = 6 Then                            'ln'

If n0 > 0 Then
panel.Caption = Log(n0)
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

ElseIf Index = 7 Then                            'log10'

If n0 > 0 Then
panel.Caption = Log(n0) / Log(10)
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

ElseIf Index = 8 Then                            '1/x'

If n0 <> 0 Then
panel.Caption = 1 / n0
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

ElseIf Index = 9 Then                            'square root'

If n0 >= 0 Then
panel.Caption = Sqr(n0)
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

ElseIf Index = 10 Then                           'cube root'

If n0 >= 0 Then
panel.Caption = n0 ^ (1 / 3)
Else
panel.Caption = "Error"
panel.FontSize = 26
End If

ElseIf Index = 11 Then                           'y root'

panel.Caption = "0"

Sop = "x^(1/y)"

panel.FontSize = 26

ElseIf Index = 12 Then                           'x^2'

panel.Caption = n0 ^ 2

ElseIf Index = 13 Then                           'x^3'

panel.Caption = n0 ^ 3

ElseIf Index = 14 Then                           'x^y'

panel.Caption = "0"

Sop = "x^y"

panel.FontSize = 26

ElseIf Index = 15 Then                           '10^x'

panel.Caption = 10 ^ n0

ElseIf Index = 16 Then                           'sinx'
 
panel.Caption = Sin(n0 * 4 * Atn(1) / 180)

ElseIf Index = 17 Then                           'cosx'
 
panel.Caption = Cos(n0 * 4 * Atn(1) / 180)
 
ElseIf Index = 18 Then                           'tanx'

panel.Caption = Tan(n0 * 4 * Atn(1) / 180)

ElseIf Index = 19 Then                           'Pi'

panel.Caption = 4 * Atn(1)

End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Dim numIndex As Integer

Select Case KeyCode

        Case vbKey0, vbKeyNumpad0:  numIndex = 0
        Case vbKey1, vbKeyNumpad1:  numIndex = 1
        Case vbKey2, vbKeyNumpad2:  numIndex = 2
        Case vbKey3, vbKeyNumpad3:  numIndex = 3
        Case vbKey4, vbKeyNumpad4:  numIndex = 4
        Case vbKey5, vbKeyNumpad5:  numIndex = 5
        Case vbKey6, vbKeyNumpad6:  numIndex = 6
        Case vbKey7, vbKeyNumpad7:  numIndex = 7
        Case vbKey8, vbKeyNumpad8:  numIndex = 8
        Case vbKey9, vbKeyNumpad9:  numIndex = 9
        
        Case Else: Exit Sub

End Select

CommandNum(numIndex).SetFocus
CommandNum_Click numIndex

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

       Case 43, vbKeyAdd: CommandAdd.SetFocus                             '+'
                          Call CommandAdd_Click
                
       Case 45, vbKeySubtract: CommandSubtract.SetFocus                   '-'
                               Call CommandSubtract_Click
                
       Case 42, vbKeyMultiply, 120, 88: CommandMultiply.SetFocus          '*,x,X'
                                        Call CommandMultiply_Click
                
       Case 47, vbKeyDivide: CommandDivide.SetFocus                       '/'
                             Call CommandDivide_Click
                             
       Case 46: CommandDecimal.SetFocus                                   '.'
                Call CommandDecimal_Click
                          
       Case 61: CommandEquals.SetFocus                                    '='
                Call CommandEquals_Click
                
       Case 99, 67: CommandClear.SetFocus                                 'c,C'
                    Call CommandClear_Click
                    
       Case 95: CommandMorP.SetFocus                                      '_'
                Call CommandMorP_Click
                
       Case 8: CommandBackSpace.SetFocus                                  '<-'
               Call CommandBackSpace_Click
               
       Case 27: CommandExit.SetFocus                                      'Esc'
                Call CommandExit_Click
                
       Case Else: Exit Sub
       
End Select

End Sub
