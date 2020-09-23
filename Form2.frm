VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   LinkTopic       =   "Form2"
   ScaleHeight     =   1110
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1305
      Top             =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "This could take a minute or so."
      Height          =   195
      Left            =   810
      TabIndex        =   2
      Top             =   855
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Please Wait ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   630
      TabIndex        =   1
      Top             =   405
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Getting Current Weather"
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
      Left            =   405
      TabIndex        =   0
      Top             =   90
      Width           =   2985
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   3795
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.MousePointer = 11
Form1.MousePointer = 11
End Sub

Private Sub Form_Resize()
Form2.Left = Form1.Left + (Form1.Width / 2) - (Form2.Width / 2)
Form2.Top = Form1.Top + (Form1.Height / 2) - (Form2.Height / 2)
End Sub

Private Sub Timer1_Timer()
If i = 2 Then
i = 0
Timer1.Enabled = False
Form1.GetWeather Form1.Text1
Me.MousePointer = 0
Form1.MousePointer = 0
Unload Me
End If
i = i + 1
End Sub
