VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ToolTip on demand"
      Height          =   405
      Left            =   4050
      TabIndex        =   5
      Top             =   4110
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3630
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   1230
      Width           =   2235
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   360
      ScaleHeight     =   4215
      ScaleWidth      =   2925
      TabIndex        =   0
      Top             =   150
      Width           =   2955
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   2610
         Width           =   2385
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTip label in Picturebox"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   210
         MouseIcon       =   "Form1.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ToolTip label on Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   3420
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   570
      Width           =   1785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TT As CTooltip
Dim m_bInLable As Boolean

Private Sub Command1_Click()
m_bInLable = True
      TT.Title = "Long message tooltip"
      TT.DelayTime = 50
      TT.VisibleTime = 12000
      TT.TipText = "(Long tooltip)" & vbCrLf & "Click the Carpet Area option, 2 boxes appear" & vbCrLf & _
      "Enter your set of measurments such as 25 x 35 pressing the enter key after each measurment" & vbCrLf & _
      "Eg; 25 Enter, 35 Enter, continue with all your carpet area measurments." & vbCrLf & vbCrLf & _
      "(stairs, hallways and living rooms)" & vbCrLf & _
      "Select each option and enter your measurments." & vbCrLf & _
       "Pressing the Enter key after each measurment" & vbCrLf & _
      "F1 for detailed help"
     
      TT.Create Command1.hwnd
End Sub

Private Sub Form_Load()
   Set TT = New CTooltip
   TT.Style = TTBalloon
   TT.Icon = TTIconInfo
   Call LoadComboTooltips
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If m_bInLable Then
      m_bInLable = False
      TT.Destroy
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DeleteComboBoxToolTips Combo1, Me.Name
    DeleteComboBoxToolTips Combo2, Me.Name
   
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Not m_bInLable Then
      m_bInLable = True
      TT.Title = "Multiline tooltip in"
      TT.TipText = "Label1" & vbCrLf & _
      "In a picturebox"
      TT.Create Picture1.hwnd
   End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_bInLable Then
      m_bInLable = True
      TT.Title = "Multiline tooltip"
      TT.TipText = "Label2 on a form"
      TT.Create Me.hwnd
   End If
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If m_bInLable Then
      m_bInLable = False
      TT.Destroy
   End If
End Sub
Sub LoadComboTooltips()
CreateComboBoxToolTip Combo1, Me.Name, "Information", "Combo Tooltip" & vbCrLf & _
   "Over Combo1 in a picturebox " & vbCrLf & "(Optional)"

CreateComboBoxToolTip Combo2, Me.Name, "Information", "Combo Tooltip" & vbCrLf & _
"Over Combo1 in a picturebox " & vbCrLf & "(Optional)"
End Sub
