VERSION 5.00
Begin VB.Form form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   Caption         =   "The Bitmap Project"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10125
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   5280
      TabIndex        =   28
      Top             =   1050
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   5280
      TabIndex        =   27
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   5280
      TabIndex        =   26
      Top             =   1860
      Width           =   550
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   360
      Picture         =   "form1.frx":0CCA
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   120
      Picture         =   "form1.frx":166C
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   24
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   120
      Picture         =   "form1.frx":200E
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   23
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   5880
      Picture         =   "form1.frx":29B0
      ScaleHeight     =   450
      ScaleWidth      =   585
      TabIndex        =   16
      Top             =   1800
      Width           =   585
   End
   Begin VB.PictureBox Picture9 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      Picture         =   "form1.frx":3232
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      Picture         =   "form1.frx":3BD4
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Picture         =   "form1.frx":4576
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2640
      Picture         =   "form1.frx":4F18
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   120
      Width           =   495
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   4440
      Picture         =   "form1.frx":58BA
      ScaleHeight     =   5445
      ScaleWidth      =   5445
      TabIndex        =   11
      Top             =   600
      Width           =   5445
   End
   Begin VB.CommandButton Command4 
      Caption         =   "menu5"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "menu4"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "menu3"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      Picture         =   "form1.frx":6DE0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   120
      Width           =   195
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      Picture         =   "form1.frx":702A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   840
      Width           =   195
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      Picture         =   "form1.frx":7274
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   600
      Width           =   195
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      Picture         =   "form1.frx":74BE
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   360
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "menu2"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the button to maximize to see all of this program."
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "The indexing starts with MENU2. MENU2 has an Index: 0, MENU3 Index: 1, MENU4 Index: 2, and MENU5 Index: 3."
      Height          =   615
      Left            =   840
      TabIndex        =   20
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "The ""Name:"", for the Caption: MENU2 is mnufile1, MENU3 is mnufile2, MENU4 is mnufile3, and MENU5 is mnufile4."
      Height          =   735
      Left            =   600
      TabIndex        =   19
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "The index box is blank at MENU1 and the name of menu1 is mnuF0Main."
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Indexing the submenu is a major thing for this function to work."
      Height          =   495
      Left            =   1080
      TabIndex        =   17
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"form1.frx":7708
      Height          =   855
      Left            =   600
      TabIndex        =   10
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Set the picture autosize to true in the properties box"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Click these buttons to add the images to the popup menu"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuF0Main 
      Caption         =   "MENU1"
      Begin VB.Menu mnufile1 
         Caption         =   "MENU2"
         Index           =   0
      End
      Begin VB.Menu mnufile2 
         Caption         =   "MENU3"
         Index           =   1
      End
      Begin VB.Menu mnufile3 
         Caption         =   "MENU4"
         Index           =   2
      End
      Begin VB.Menu mnufile4 
         Caption         =   "MENU5"
         Index           =   3
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ALL THE WORK IS IN THE MOD!~
' RUN THE PROGRAM LOOK AT THE MENU AND CLICK THE BUTTONS
'IT'S PRETTY SIMPLE
'HOPE IT'S WHAT YER LOOKIN FER!
'JIMIDOTCOM@YAHOO.COM

Private Sub Command1_Click()
Call pictopop1
End Sub

Private Sub Command2_Click()
Call pictopop2
End Sub

Private Sub Command3_Click()
Call pictopop3
End Sub

Private Sub Command4_Click()
Call pictopop4
End Sub


Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
'The stuff below is just for the demo
form1.Text1 = ""
form1.Text2 = "mnuF0Main"
form1.Text3 = "MENU1"
End Sub


Private Sub mnufile1_Click(Index As Integer)
Text1 = "0"
Text2 = "mnufile1"
Text3 = "MENU2"
Text1.SetFocus
'The stuff above is just for the demo
MsgBox "This is where you can add a function!  Like , form2.show"

End Sub

Private Sub mnufile2_Click(Index As Integer)
Text1 = "1"
Text2 = "mnufile2"
Text3 = "MENU3"
Text1.SetFocus
'The stuff above is just for the demo
MsgBox "This is where you can add a function!  Like , form2.show"

End Sub

Private Sub mnufile3_Click(Index As Integer)
Text1 = "2"
Text2 = "mnufile3"
Text3 = "MENU4"
Text1.SetFocus
'The stuff above is just for the demo
MsgBox "This is where you can add a function!  Like , form2.show"

End Sub

Private Sub mnufile4_Click(Index As Integer)
Text1 = "3"
Text2 = "mnufile4"
Text3 = "MENU5"
Text1.SetFocus
'The stuff above is just for the demo
MsgBox "This is where you can add a function!  Like , form2.show"

End Sub
