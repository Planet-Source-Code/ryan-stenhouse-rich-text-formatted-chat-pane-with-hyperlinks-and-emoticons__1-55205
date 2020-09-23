VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.UserControl nlRTB 
   BackStyle       =   0  'Transparent
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   ScaleHeight     =   6945
   ScaleWidth      =   7560
   Begin RichTextLib.RichTextBox rtbMain 
      Height          =   1335
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"nlRTB.ctx":0000
      MouseIcon       =   "nlRTB.ctx":007B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imglEmotes 
      Left            =   6705
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":01DD
            Key             =   ""
            Object.Tag             =   ":D :d :-d :-D"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":0807
            Key             =   ""
            Object.Tag             =   ":@ :-@"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":0E31
            Key             =   ""
            Object.Tag             =   ":) :-)"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":145B
            Key             =   ""
            Object.Tag             =   ":'("
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":1A85
            Key             =   ""
            Object.Tag             =   ":S :s :-s :-S"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":20AF
            Key             =   ""
            Object.Tag             =   ":| :-| :-\"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":26D9
            Key             =   ""
            Object.Tag             =   "+o("
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":2D03
            Key             =   ""
            Object.Tag             =   ":( :-("
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":332D
            Key             =   ""
            Object.Tag             =   "^-)"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":3957
            Key             =   ""
            Object.Tag             =   ":-< :< :[ :-{"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":3F81
            Key             =   ""
            Object.Tag             =   ";) ;-)"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":45AB
            Key             =   ""
            Object.Tag             =   "XD xd Xd xD"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":4BD5
            Key             =   ""
            Object.Tag             =   ":P :-P :p :-p"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlRTB.ctx":51FF
            Key             =   ""
            Object.Tag             =   ":o :-o :O :-O"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu Devider 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As"
      End
   End
End
Attribute VB_Name = "nlRTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Actions As Long
Private Thoughts As Long
Private Emotes As Boolean
Private sLinks As Boolean




Private Type STRUCTlinks
  beginC As Long
  endC As Long
  URL As String
End Type

Private Type POINTL
  x As Long
  y As Long
End Type

Const EM_CHARFROMPOS = &HD7
Const WM_PASTE = &H302
Const SW_NORMAL = 1
 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long




Private Links() As STRUCTlinks


Public Property Get ActionColour() As Long
ActionColour = Actions
End Property

Public Property Let ActionColour(ByVal NewColour As Long)
Actions = NewColour
PropertyChanged "ActionColour"
End Property

Public Property Get ThoughtColour() As Long
ThoughtColour = Thoughts
End Property

Public Property Let ThoughtColour(ByVal NewColour As Long)
Thoughts = NewColour
PropertyChanged "ThoughtColour"
End Property

Public Property Get UseEmotes() As Boolean
UseEmotes = Emotes
End Property

Public Property Let UseEmotes(ByVal value As Boolean)
Emotes = value
PropertyChanged "UseEmotes"
End Property

Public Property Get ShowURLs() As Boolean
ShowURLs = sLinks
End Property

Public Property Let ShowURLs(ByVal value As Boolean)
sLinks = value
PropertyChanged "ShowURLs"
End Property


Private Sub mnuClear_Click()
    
    ClearScreen
    ReDim Links(0)
    
End Sub



Private Sub rtbMain_Change()
rtbMain.SelStart = Len(rtbMain.Text)
End Sub


Private Sub rtbMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    
    PopupMenu mnuPopup
       rtbMain.SelStart = Len(rtbMain.Text)
       
Else
     Dim pt As POINTL, charpos As Long, i As Long
     
     rtbMain.Locked = False
     
  pt.x = x \ Screen.TwipsPerPixelX
  pt.y = y \ Screen.TwipsPerPixelY
  'get the character closest to the mouse cursor
  charpos = SendMessage(rtbMain.hwnd, EM_CHARFROMPOS, 0, pt)
  
  rtbMain.Locked = True
  
  rtbMain.MousePointer = rtfDefault
  'see if the character is in any links-
  'start with the last link entered because it's more likely to be clicked
  For i = UBound(Links) To 1 Step -1
    If charpos > Links(i).beginC And charpos < Links(i).endC Then
  
  Dim link As Variant
    
    link = ShellExecute(hwnd, "Open", Links(i).URL, &O0, &O0, SW_NORMAL)
  
      Exit For  'found our link, don't need to look anymore
    End If
  Next i
    rtbMain.SelStart = Len(rtbMain.Text)
    
End If

rtbMain.SelUnderline = False

End Sub

Private Sub rtbMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pt As POINTL, charpos As Long, i As Long
  pt.x = x \ Screen.TwipsPerPixelX
  pt.y = y \ Screen.TwipsPerPixelY
  'get the character closest to the mouse cursor
  
  rtbMain.Locked = False
  
  charpos = SendMessage(rtbMain.hwnd, EM_CHARFROMPOS, 0, pt)
  
  rtbMain.Locked = True
  
  rtbMain.MousePointer = rtfDefault
  'see if the character is in any links-
  'start with the last link entered because it's more likely to be clicked
  For i = UBound(Links) To 1 Step -1
    If charpos > Links(i).beginC And charpos < Links(i).endC Then
      rtbMain.MousePointer = rtfCustom
      Exit For  'found our link, don't need to look anymore
    End If
  Next i
End Sub

Private Sub UserControl_Initialize()
  ReDim Links(0)
  Action = &H800080
  Thoughts = vbBlue
  Emotes = True
  sLinks = True
End Sub

Private Sub UserControl_InitProperties()
    
'Enable everything by default
    ActionColour = &H800080
    ThoughtColour = vbBlue
    ShowURLs = True
    UseEmotes = True

End Sub

Private Sub UserControl_Resize()


On Error Resume Next

    rtbMain.Left = UserControl.ScaleLeft
    rtbMain.Top = UserControl.ScaleTop
    rtbMain.Width = UserControl.ScaleWidth
    rtbMain.Height = UserControl.ScaleHeight
    

End Sub

Public Sub UserSpeaks(UserName As String, Message As String, Bold As Boolean, Italic As Boolean, Underline As Boolean, Colour As Long, Font As String)
    
On Error Resume Next
    
    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, iCC As Integer
    Dim CharCombo() As String
    Dim ClipboardContents As Variant
    Dim bClipHasImage As Boolean
    
    Dim linkStart As Long
    Dim linkEnd As Long
    Dim URL As String
        
    Dim startpoint As Long
    Dim endpoint As Long
    
Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")


    
    With rtbMain
        
        .Locked = False
                
        .SelStart = Len(.Text)
        
        startpoint = .SelStart
        
        .SelBold = False
        .SelColor = RGB(85, 0, 149)
        .SelText = "    " + UserName + " :  "
        .SelBold = False
        .SelColor = Colour
        .SelBold = Bold
        .SelItalic = Italic
        .SelUnderline = Underline
        .SelFontName = Font
        
            
       lStartMessage = .SelStart - 1  'Where the new message begns i(search starts here
                                        '   for the icons)
        
        .SelText = Message + vbCrLf
        .SelColor = vbBlack
        
       endpoint = Len(.Text)
      
        
    End With
    
 If Emotes = True Then
    
    For i = 1 To imglEmotes.ListImages.Count
    
        'Get all the keystroke combos needed to make the emote
        CharCombo = Split(imglEmotes.ListImages(i).Tag, " ")
        
        For j = 0 To UBound(CharCombo)
            
            lImagePos = InStr(lStartMessage, rtbMain.Text, CharCombo(j))
            
            While lImagePos > 0
                
                rtbMain.SelStart = lImagePos - 1
                rtbMain.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                rtbMain.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData imglEmotes.ListImages(i).Picture        'Set the icon in it
                rtbMain.SelStart = lImagePos - 1
                SendMessage rtbMain.hwnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                lImagePos = InStr(lImagePos, rtbMain.Text, CharCombo(j))
                
            Wend
            
            
        Next j
        
    
  Next i
  
  End If
  
  linkStart = InStr(lStartMessage, rtbMain.Text, "http://")
  
 
 If linkStart <> 0 Then
   If sLinks = True Then
              While linkStart > 0
                
                linkEnd = InStr(linkStart, rtbMain.Text, " ")
                If linkEnd = 0 Then linkEnd = Len(rtbMain.Text)
                 
                With rtbMain
                    
                    .SelStart = linkStart - 1
                    .SelLength = linkEnd - linkStart
                    
                    URL = .SelText
                    
                    .SelColor = vbBlue
                    .SelUnderline = True
                    
                    
                    ReDim Preserve Links(UBound(Links) + 1)
                    Links(UBound(Links)).beginC = linkStart
                    Links(UBound(Links)).endC = linkEnd
                    Links(UBound(Links)).URL = URL
                           
                       
                End With
                    
                linkStart = InStr(linkEnd, rtbMain.Text, "http://")
                
              Wend
              
            rtbMain.SelStart = Len(rtbMain.Text)
            rtbMain.Locked = True
            Clipboard.Clear
            
            rtbMain.SelUnderline = False
    Else
     
         With rtbMain
               
            .SelStart = startpoint
            .SelLength = (endpoint - startpoint)
            .SelText = ""
            .SelStart = startpoint
            .SelColor = &H808080
            .SelFontName = "Tahoma"
            .SelText = "        "
            .SelFontName = "Marlett"
            .SelText = "8"
            .SelFontName = "Tahoma"
            .SelText = "A message from " & UserName & " could not be displayed, as it contained a URL. If you wish to view URLs, please change your chat settings." + vbCrLf
            .SelColor = 0
            
        End With
        
    End If
    
  End If
  
End Sub

Public Sub YouSpeak(UserName As String, Message As String, Bold As Boolean, Italic As Boolean, Underline As Boolean, Colour As Long, Font As String)
      
On Error Resume Next

        Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, iCC As Integer
    Dim CharCombo() As String
    Dim ClipboardContents As Variant
    Dim bClipHasImage As Boolean
    
    Dim linkStart As Long
    Dim linkEnd As Long
    Dim URL As String
Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")


    
    With rtbMain
        
        .Locked = False
        
        .SelStart = Len(.Text)
        .SelBold = True
        .SelColor = vbBlack
        .SelText = "    " + UserName + " :  "
        .SelBold = False
        .SelColor = Colour
        .SelBold = Bold
        .SelItalic = Italic
        .SelUnderline = Underline
        .SelFontName = Font
        
            
       lStartMessage = .SelStart
       
        .SelText = Message + vbCrLf
        .SelColor = vbBlack
     
    End With
    
  
  If Emotes = True Then
  
  For i = 1 To imglEmotes.ListImages.Count
    
        'Get all the keystroke combos needed to make the emote
        CharCombo = Split(imglEmotes.ListImages(i).Tag, " ")
        
        For j = 0 To UBound(CharCombo)
            
            lImagePos = InStr(lStartMessage, rtbMain.Text, CharCombo(j))
            
            While lImagePos > 0
                
                rtbMain.SelStart = lImagePos - 1
                rtbMain.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                rtbMain.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData imglEmotes.ListImages(i).Picture        'Set the icon in it
                rtbMain.SelStart = lImagePos - 1
                SendMessage rtbMain.hwnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                lImagePos = InStr(lImagePos, rtbMain.Text, CharCombo(j))
                
            Wend
            
            
        Next j
        
    
  Next i
  
  End If
  
   linkStart = InStr(lStartMessage, rtbMain.Text, "http://")
  
  While linkStart > 0
    
    linkEnd = InStr(linkStart, rtbMain.Text, " ")
    
     If linkEnd = 0 Then linkEnd = Len(rtbMain.Text)
    
    With rtbMain
        
        .SelStart = linkStart - 1
        .SelLength = linkEnd - linkStart
        
        URL = .SelText
        
        .SelColor = vbBlue
        .SelUnderline = True
        
        
        ReDim Preserve Links(UBound(Links) + 1)
        Links(UBound(Links)).beginC = linkStart
        Links(UBound(Links)).endC = linkEnd
        Links(UBound(Links)).URL = URL
               
           
    End With
        
    linkStart = InStr(linkEnd, rtbMain.Text, "http://")
    
  Wend
  
  rtbMain.SelStart = Len(rtbMain.Text)
  rtbMain.Locked = True
  Clipboard.Clear
  
  rtbMain.SelUnderline = False
  
End Sub


Public Sub UserActions(UserName As String, Message As String)
     
On Error Resume Next
    
    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, j As Integer
    Dim CharCombo() As String

    
Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")
    
    With rtbMain
        
                
        .SelStart = Len(.Text)
        .SelBold = False
        .SelUnderline = False
        .SelItalic = True
        .SelColor = Action
        .SelText = "    " + UserName + " "
        
        lStartMessage = .SelStart
        .Locked = False
        
        .SelText = Message + vbCrLf
        .SelColor = vbBlack
        .SelItalic = False
        .SelStart = Len(.Text)
        
        
    End With
    
 If Emotes = True Then
    
  
  For i = 1 To imglEmotes.ListImages.Count
    
        'Get all the keystroke combos needed to make the emote
        CharCombo = Split(imglEmotes.ListImages(i).Tag, " ")
        
        For j = 0 To UBound(CharCombo)
            
            lImagePos = InStr(lStartMessage, rtbMain.Text, CharCombo(j))
            
            While lImagePos > 0
                
                rtbMain.SelStart = lImagePos - 1
                rtbMain.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                rtbMain.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData imglEmotes.ListImages(i).Picture        'Set the icon in it
                rtbMain.SelStart = lImagePos - 1
                SendMessage rtbMain.hwnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                lImagePos = InStr(lImagePos, rtbMain.Text, CharCombo(j))
                
            Wend
            
            
        Next j
        
    
  Next i
  
  rtbMain.SelStart = Len(rtbMain.Text)
  rtbMain.Locked = True
  Clipboard.Clear
   
 End If
  
  rtbMain.SelUnderline = False

End Sub

Public Sub UserThinks(UserName As String, Message As String)
        
On Error Resume Next
   
    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, j As Integer
    Dim CharCombo() As String
 
    
Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")
    
    With rtbMain
        
        .SelStart = Len(.Text)
        .SelBold = False
        .SelUnderline = False
        .SelItalic = True
        .SelColor = Thoughts
        
        .SelText = "    " + UserName + ": ~"
        
        lStartMessage = .SelStart
        .Locked = False
        
        .SelText = Message + "~" + vbCrLf
        .SelColor = vbBlack
        .SelItalic = False
        .SelStart = Len(.Text)
        
        
    End With
      
 If Emotes = True Then
  
  For i = 1 To imglEmotes.ListImages.Count
    
        'Get all the keystroke combos needed to make the emote
        CharCombo = Split(imglEmotes.ListImages(i).Tag, " ")
        
        For j = 0 To UBound(CharCombo)
            
            lImagePos = InStr(lStartMessage, rtbMain.Text, CharCombo(j))
            
            While lImagePos > 0
                
                rtbMain.SelStart = lImagePos - 1
                rtbMain.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                rtbMain.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData imglEmotes.ListImages(i).Picture        'Set the icon in it
                rtbMain.SelStart = lImagePos - 1
                SendMessage rtbMain.hwnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                lImagePos = InStr(lImagePos, rtbMain.Text, CharCombo(j))
                
            Wend
            
            
        Next j
        
    
  Next i
  
  rtbMain.SelStart = Len(rtbMain.Text)
  rtbMain.Locked = True
  Clipboard.Clear
  
  End If
  
  rtbMain.SelUnderline = False
  
End Sub

Public Sub UserOOCS(UserName As String, Message As String)
        
On Error Resume Next

    Dim linkStart As Long
    Dim linkEnd As Long
    Dim URL As String

    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, j As Integer
    Dim CharCombo() As String
    
    Dim startpoint As Long
    Dim endpoint As Long
    
    
    
  Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")
  
  With rtbMain
        
        .SelStart = Len(.Text)
        
        startpoint = .SelStart
        
        .SelBold = True
        .SelColor = vbBlack
        .SelText = "    " + UserName + " : "
        .SelBold = False
        
            
        lStartMessage = .SelStart
        .Locked = False
        
             
        .SelText = "((" + Message + "))" + vbCrLf
        .SelStart = Len(.Text)
        
        endpoint = .SelStart
        
    End With
        
  If Emotes = True Then
  
  For i = 1 To imglEmotes.ListImages.Count
    
        'Get all the keystroke combos needed to make the emote
        CharCombo = Split(imglEmotes.ListImages(i).Tag, " ")
        
        For j = 0 To UBound(CharCombo)
            
            lImagePos = InStr(lStartMessage, rtbMain.Text, CharCombo(j))
            
            While lImagePos > 0
                
                rtbMain.SelStart = lImagePos - 1
                rtbMain.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                rtbMain.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData imglEmotes.ListImages(i).Picture        'Set the icon in it
                rtbMain.SelStart = lImagePos - 1
                SendMessage rtbMain.hwnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                lImagePos = InStr(lImagePos, rtbMain.Text, CharCombo(j))
                
            Wend
            
            
        Next j
        
    
  Next i
  
  End If
  
     
    linkStart = InStr(lStartMessage, rtbMain.Text, "http://")
  
 
 If linkStart <> 0 Then
   If sLinks = True Then
              While linkStart > 0
                
                linkEnd = InStr(linkStart, rtbMain.Text, " ")
                If linkEnd = 0 Then linkEnd = Len(rtbMain.Text)
                 
                With rtbMain
                    
                    .SelStart = linkStart - 1
                    .SelLength = linkEnd - linkStart
                    
                    URL = .SelText
                    
                    .SelColor = vbBlue
                    .SelUnderline = True
                    
                    
                    ReDim Preserve Links(UBound(Links) + 1)
                    Links(UBound(Links)).beginC = linkStart
                    Links(UBound(Links)).endC = linkEnd
                    Links(UBound(Links)).URL = URL
                           
                       
                End With
                    
                linkStart = InStr(linkEnd, rtbMain.Text, "http://")
                
              Wend
              
            rtbMain.SelStart = Len(rtbMain.Text)
            rtbMain.Locked = True
            Clipboard.Clear
            
            rtbMain.SelUnderline = False
    Else
     
         With rtbMain
               
            .SelStart = startpoint
            .SelLength = (endpoint - startpoint)
            .SelText = ""
            .SelStart = startpoint
            .SelColor = &H808080
            .SelFontName = "Tahoma"
            .SelText = "        "
            .SelFontName = "Marlett"
            .SelText = "8"
            .SelFontName = "Tahoma"
            .SelText = "A message from " & UserName & " could not be displayed, as it contained a URL. If you wish to view URLs, please change your chat settings." + vbCrLf
            .SelColor = 0
            
        End With
        
    End If
    
  End If
  
End Sub


Public Sub RoomAction(Message As String)
    
Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")
    
    With rtbMain
          .SelStart = Len(.Text)
        .SelColor = &H808080
        .SelFontName = "Tahoma"
        .SelText = "        "
        .SelFontName = "Marlett"
        .SelText = "8"
        .SelFontName = "Tahoma"
        .SelText = Message + vbCrLf
        .SelColor = 0
        
    End With
End Sub

Public Sub Connecting(Server As String)
    
    With rtbMain
          .SelStart = Len(.Text)
        .SelColor = RGB(0, 128, 0)
        

        .SelFontName = "Tahoma"
        .SelText = "Connecting to " & Server & "... " + vbCrLf
        .SelColor = 0
        
    End With
End Sub

Public Sub ServerMessage(Message As String)
        
On Error Resume Next

    Dim linkStart As Long
    Dim linkEnd As Long
    Dim URL As String
    
Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")
    
    With rtbMain
          .SelStart = Len(.Text)
        .SelColor = vbRed

        .SelFontName = "Tahoma"
        .SelText = "Broadcast Message from Server: " + vbCrLf
        .SelColor = 0
        .SelText = Message + vbCrLf
        
    End With
End Sub

Public Sub Connected()
    
    With rtbMain
          .SelStart = Len(.Text)
        .SelColor = vbRed

        .SelFontName = "Tahoma"
        .SelText = "Connected! " + vbCrLf + vbCrLf
        .SelColor = 0
        
    End With
End Sub

Public Sub KickOrBan(Kicker As String, Kickee As String, Message As String)
  
 Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")
 
  With rtbMain
        
        .SelStart = Len(.Text)
        .SelColor = vbRed
        .SelBold = True
        .SelFontName = "Tahoma"
        .SelText = "    " + "Host " + Kicker + " has kicked " + Kickee + " out of the chatroom: " + Message + vbCrLf
        .SelColor = 0
        .SelBold = False
        
    End With
End Sub

Public Sub Notice(From As String, Message As String)
  
 Message = Replace(Replace(Replace(Message, vbCrLf, ""), vbLf, ""), vbCr, "")
  
  With rtbMain
        
        .SelStart = Len(.Text)
        .SelColor = 0
        .SelBold = True
        .SelFontName = "Tahoma"
        .SelText = "    " + "Message from " + From + ":" + vbCrLf
        .SelBold = False
        .SelText = "    " + Message & vbCrLf
        .SelColor = 0
        .SelBold = False
        
    End With
End Sub

Public Sub Topic(Topic As String)

    
On Error Resume Next

    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, j As Integer
    Dim CharCombo() As String

    
    Dim linkStart As Long
    Dim linkEnd As Long
    Dim URL As String

Topic = Replace(Replace(Replace(Topic, vbCrLf, ""), vbLf, ""), vbCr, "")

With rtbMain
    
    .Locked = False
    
    .SelStart = Len(.Text)
    .SelColor = &H958055
    .SelBold = False
    .SelFontName = "Tahoma"
    .SelText = "This Chatroom's Topic Is: "
    .SelBold = False
    .SelColor = 0
    
       lStartMessage = Len(.Text) - 1  'Where the new message begns i(search starts here
                                        '   for the icons)
    
    .SelText = Topic + vbCrLf + vbCrLf
    
End With

If Emotes = True Then
        
    For i = 1 To imglEmotes.ListImages.Count
    
        'Get all the keystroke combos needed to make the emote
        CharCombo = Split(imglEmotes.ListImages(i).Tag, " ")
        
        For j = 0 To UBound(CharCombo)
            
            lImagePos = InStr(lStartMessage, rtbMain.Text, CharCombo(j))
            
            While lImagePos > 0
                
                rtbMain.SelStart = lImagePos - 1
                rtbMain.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                rtbMain.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData imglEmotes.ListImages(i).Picture        'Set the icon in it
                rtbMain.SelStart = lImagePos - 1
                SendMessage rtbMain.hwnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                lImagePos = InStr(lImagePos, rtbMain.Text, CharCombo(j))
                
            Wend
            
            
        Next j
        
    
  Next i
  
 End If
 

   linkStart = InStr(lStartMessage, rtbMain.Text, "http://")
  
  While linkStart > 0
    
    linkEnd = InStr(linkStart, rtbMain.Text, " ")
     If linkEnd = 0 Then linkEnd = Len(rtbMain.Text)
     
    With rtbMain
        
        .SelStart = linkStart - 1
        .SelLength = linkEnd - linkStart
        
        URL = .SelText
        
        .SelColor = vbBlue
        .SelUnderline = True
        
        
        ReDim Preserve Links(UBound(Links) + 1)
        Links(UBound(Links)).beginC = linkStart
        Links(UBound(Links)).endC = linkEnd
        Links(UBound(Links)).URL = URL
               
           
    End With
        
    linkStart = InStr(linkEnd, rtbMain.Text, "http://")
    
  Wend
  
  rtbMain.SelStart = Len(rtbMain.Text)
  rtbMain.Locked = True
  Clipboard.Clear

  rtbMain.SelUnderline = False

End Sub

Public Sub ClearScreen()
    
    Dim x As Variant
    
    x = MsgBox("This will clear the chat window, are you sure?", vbYesNo + vbQuestion, "Really Clear?")
    
    If x = vbYes Then
        
        rtbMain.Text = ""
        
    End If
    
End Sub


