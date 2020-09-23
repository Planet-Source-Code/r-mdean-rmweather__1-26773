VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "RM Weather"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Go"
      Height          =   285
      Left            =   5895
      TabIndex        =   5
      Top             =   90
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4635
      TabIndex        =   4
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "E&xit"
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Top             =   3195
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "Form1.frx":0CCA
      ScaleHeight     =   465
      ScaleWidth      =   6375
      TabIndex        =   6
      Top             =   0
      Width           =   6405
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Zip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3645
         TabIndex        =   7
         Top             =   90
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   135
      ScaleHeight     =   780
      ScaleWidth      =   780
      TabIndex        =   8
      Top             =   1305
      Width           =   780
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001 Richard Dean. All Rights Reserved"
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   3285
      Width           =   3765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3105
      TabIndex        =   13
      Top             =   1530
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1125
      TabIndex        =   12
      Top             =   2070
      Width           =   60
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1125
      TabIndex        =   11
      Top             =   1530
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1215
      TabIndex        =   10
      Top             =   2205
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   315
      TabIndex        =   9
      Top             =   2070
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   495
      Width           =   75
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   6315
   End
   Begin VB.Shape Shape2 
      Height          =   3570
      Left            =   0
      Top             =   0
      Width           =   6405
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   135
      TabIndex        =   3
      Top             =   765
      Width           =   6120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show 1, Form1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Form2.Show 1, Form1
End Sub

Public Sub GetWeather(sZip As String)
Dim S3W As S3Weather.Current
Set S3W = New S3Weather.Current
S3W.GetWeather sZip
S3W.BlockTracking = True
Label3.Caption = S3W.City & " " & S3W.State & " (" & Text1 & ")"
Label4.Caption = S3W.Reported
Picture2.Picture = LoadPicture(App.Path & "\images\" & S3W.ImageID & ".gif")
Label1.Caption = S3W.Temperature
Label6.Caption = Chr(176) & "F"
Label6.Left = Label1.Width + 300
Label6.Top = Label1.Top + 50
Label7.Caption = S3W.Forecast
Label8.Caption = "Feels like: " & Replace(S3W.FeelsLike, "&deg;", Chr(176))
Label9.Caption = "Wind: " & S3W.Wind & vbCrLf & _
"Dew point: " & Replace(S3W.DewPoint, "&deg;", Chr(176)) & vbCrLf & "Humidity: " & S3W.Humidity & vbCrLf & _
"Visibality: " & S3W.Version & vbCrLf & "Barometer: " & S3W.Barometer
Set S3W = Nothing
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
formdrag Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Command3_Click
End If
End Sub
