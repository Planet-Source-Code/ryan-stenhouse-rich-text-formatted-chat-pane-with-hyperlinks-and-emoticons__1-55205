VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Testing Ryan Stenhouse's Chat Log"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Eixit 
      Appearance      =   0  'Flat
      Caption         =   "Exit Demo Application"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   5895
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pre-Defined Demos"
      Height          =   735
      Left            =   45
      TabIndex        =   15
      Top             =   5085
      Width           =   6045
      Begin VB.CommandButton Demo2 
         Caption         =   "Demo Two"
         Height          =   420
         Left            =   2205
         TabIndex        =   18
         Top             =   225
         Width           =   1950
      End
      Begin VB.CommandButton Demo1 
         Caption         =   "Demo One"
         Height          =   420
         Left            =   135
         TabIndex        =   17
         Top             =   225
         Width           =   1950
      End
   End
   Begin VB.OptionButton optUnderline 
      Caption         =   "Underline Text"
      Height          =   240
      Left            =   2295
      TabIndex        =   8
      Top             =   4590
      Width           =   1365
   End
   Begin VB.OptionButton optItalic 
      Caption         =   "Italic Text"
      Height          =   240
      Left            =   1170
      TabIndex        =   7
      Top             =   4590
      Width           =   1410
   End
   Begin VB.Frame fraFontSettings 
      Caption         =   "Test Font Settings"
      Height          =   1050
      Left            =   45
      TabIndex        =   4
      Top             =   3960
      Width           =   6000
      Begin VB.OptionButton optBold 
         Caption         =   "Bold Text"
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   630
         Width           =   1365
      End
      Begin VB.ComboBox cboFonts 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Text            =   "Comic Sans MS"
         Top             =   225
         Width           =   5865
      End
      Begin VB.Label Swatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5490
         TabIndex        =   14
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Swatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5130
         TabIndex        =   13
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Swatch 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4770
         TabIndex        =   12
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Swatch 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   4410
         TabIndex        =   11
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Swatch 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4050
         TabIndex        =   10
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Swatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   3690
         TabIndex        =   9
         Top             =   630
         Width           =   330
      End
   End
   Begin VB.TextBox txtToSend 
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Text            =   "This is a sample, With :D :@ Emotes and http://www.pscode.com  Hyperlinks. Uses Windows API."
      Top             =   3510
      Width           =   4605
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Send Text"
      Height          =   375
      Left            =   4725
      TabIndex        =   2
      Top             =   3465
      Width           =   1365
   End
   Begin VB.Frame fraPane 
      Caption         =   "The Chat Pane"
      Height          =   3300
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6090
      Begin prjTestChatBox.nlRTB nlRTB1 
         Height          =   2985
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   5265
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelColor As Long

Private Sub cmdTest_Click()
Me.nlRTB1.YouSpeak "TestClient", txtToSend.Text, optBold.value, optItalic.value, optUnderline.value, SelColor, cboFonts.Text
End Sub

Private Sub Demo1_Click()
nlRTB1.Connecting "server.host.net"
nlRTB1.Connected
nlRTB1.Topic "This is an example of the Neoline Chat Frame Control. By Ryan Stenhouse :O http://www.neolinesw.co.uk"
End Sub

Private Sub Demo2_Click()
With nlRTB1
    
    .RoomAction "BobDole as joined the conversation"
    .UserSpeaks "BobDole", "Wow, this chat thing is cool", False, False, False, vbRed, "Times New Roman"
    .KickOrBan "TestClient", "BobDole", "It sure is, unlike Bob Dole!"
    
    
End With
End Sub

Private Sub Eixit_Click()
Unload Me
End Sub

Private Sub Form_Load()

'Set the default Colour
SelColor = vbBlack


    Dim X As Integer
    
    For X = 1 To Screen.FontCount
        
        'Fill the combo box up with the fonts
        cboFonts.AddItem Screen.Fonts(X)
        
    Next X
    
End Sub

Private Sub Swatch_Click(Index As Integer)
SelColor = Swatch(Index).BackColor
End Sub
