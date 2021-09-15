VERSION 5.00
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kelompok 15 - 10 Manipulasi Data String"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9902.41
   ScaleMode       =   0  'User
   ScaleWidth      =   16965.52
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Output_text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   27
      Top             =   3240
      Width           =   12975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000003&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3480
      MousePointer    =   12  'No Drop
      TabIndex        =   25
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   7320
      TabIndex        =   22
      Text            =   "0"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Reset_Command 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   21
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   20
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   19
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Lowercase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   18
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "jumlah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   17
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11760
      TabIndex        =   16
      Top             =   7560
      Width           =   4815
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11760
      TabIndex        =   15
      Top             =   6720
      Width           =   4815
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11760
      TabIndex        =   14
      Top             =   5880
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11760
      TabIndex        =   13
      Top             =   5040
      Width           =   4815
   End
   Begin VB.CommandButton Pembalikan 
      Caption         =   "Pembalikan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   12
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11760
      TabIndex        =   11
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3480
      TabIndex        =   10
      Top             =   6720
      Width           =   4815
   End
   Begin VB.TextBox Main_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Text            =   "kalimat"
      Top             =   2040
      Width           =   12975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "penggabungan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Replace String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mid String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Uppercase String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3480
      TabIndex        =   4
      Top             =   7560
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   4920
      TabIndex        =   3
      Text            =   "0"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Merge String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3480
      TabIndex        =   1
      Top             =   4200
      Width           =   4815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Output Kalimat "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   28
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Masukan Kalimat "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   26
      Top             =   1560
      Width           =   2130
   End
   Begin VB.Label Mid_Label 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   6360
      TabIndex        =   24
      Top             =   6000
      Width           =   645
   End
   Begin VB.Label Mid_Label 
      AutoSize        =   -1  'True
      Caption         =   "Awal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   3960
      TabIndex        =   23
      Top             =   6000
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "10 Manipulasi Data String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Output_text = Main_Text + Text1
End Sub

Private Sub Command10_Click()
Output_text = Right(Main_Text, 3)
End Sub

Private Sub Command2_Click()
Output_text = UCase(Main_Text)
End Sub

Private Sub Command3_Click()
Dim Awal As Integer
Dim n As Integer
Awal = Text3(0)
n = Text3(1)
Output_text = Mid(Main_Text, Awal, n)
End Sub

Private Sub Command4_Click()
Output_text = Text4
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command5_Click()
Output_text = String(10, (Main_Text))

End Sub

Private Sub Command7_Click()
Output_text = Len(Main_Text)
End Sub

Private Sub Command8_Click()
Output_text = LCase(Main_Text)
End Sub

Private Sub Command9_Click()
Output_text = Left(Main_Text, 3)
End Sub

Private Sub Pembalikan_Click()
Output_text = StrReverse(Main_Text)
End Sub

Private Sub Reset_Command_Click()
Main_Text = ""
End Sub

