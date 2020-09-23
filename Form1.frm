VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PIC BROWSER by Jongmin Baek"
   ClientHeight    =   8130
   ClientLeft      =   3765
   ClientTop       =   2055
   ClientWidth     =   9465
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   9465
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "Form1.frx":0442
      Top             =   7560
      Width           =   9255
   End
   Begin VB.VScrollBar IconScroll 
      Height          =   7525
      Left            =   9120
      Max             =   0
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   7550
      Left            =   7920
      ScaleHeight     =   7485
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   0
      Width           =   1215
      Begin VB.Image IconList 
         Height          =   1095
         Index           =   6
         Left            =   0
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Image IconList 
         Height          =   1095
         Index           =   5
         Left            =   0
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Image IconList 
         Height          =   1095
         Index           =   4
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Image IconList 
         Height          =   1095
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Image IconList 
         Height          =   1095
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Image IconList 
         Height          =   1095
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image IconList 
         Height          =   1095
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox NewImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2670
      Left            =   2040
      MousePointer    =   5  'Size
      ScaleHeight     =   2670
      ScaleWidth      =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy image to the Clipboard"
      Enabled         =   0   'False
      Height          =   325
      Left            =   1920
      TabIndex        =   10
      Top             =   7200
      Width           =   2415
   End
   Begin VB.HScrollBar HMove 
      Height          =   255
      Left            =   2040
      Max             =   0
      TabIndex        =   9
      Top             =   6960
      Width           =   5535
   End
   Begin VB.VScrollBar VMove 
      Height          =   6135
      Left            =   7560
      Max             =   0
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Sample 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   2040
      ScaleHeight     =   6135
      ScaleWidth      =   5535
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.HScrollBar ScaleBar 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2880
      Max             =   500
      Min             =   30
      TabIndex        =   3
      Top             =   480
      Value           =   100
      Width           =   4695
   End
   Begin VB.FileListBox FileList 
      Height          =   3210
      Left            =   120
      Pattern         =   "*.jpg;*.bmp;*.gif"
      TabIndex        =   2
      Top             =   4320
      Width           =   1815
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.DirListBox Folder 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Scale------>"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   480
      Width           =   855
   End
   Begin VB.Label ScaleMeter 
      Caption         =   "Scale : 100%"
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   570
      Left            =   4320
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label PathLabel 
      Caption         =   "File Path : "
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   4215
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu CommandExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu CommandCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu CommandPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu CommandCut 
         Caption         =   "&Cut"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************
'* Programmed by Jongmin Baek *
'******************************
'Please let me know if you are going to use this code in your application.
'(But this code is not hard to understand.)
'::>chunjaemanse3@netzero.net
Dim IconStart As Integer
Dim DragX As Double
Dim DragY As Double
Dim Dragnow As Boolean
Dim WWW As Double
Dim HHH As Double
Dim FileExt As String
Dim Kill_or_Not As Boolean
Dim ClipboardPath As String
Dim DefaultDrive As String
Dim DefaultPath As String
Dim FilePath As String
Dim Picture_Shown As Boolean
Dim NewImage_Width As Double
Dim NewImage_Height As Double
Private Sub Command2_Click()
Clipboard.SetData Sample, 2
End Sub
Private Sub CommandExit_Click()
End
End Sub
Private Sub Drive_Change()
c$ = Drive
On Error GoTo 10
Folder.Path = Drive
DefaultDrive = Drive
DefaultPath = Folder.Path
Exit Sub
10 Drive = DefaultDrive
Folder.Path = DefaultPath
MsgBox "Drive Reading Error", vbOKOnly, "Error!!!"
End Sub
Private Sub CommandCopy_Click()
ClipboardPath = FilePath
CommandPaste.Enabled = True
Kill_or_Not = False
End Sub
Private Sub CommandCut_Click()
ClipboardPath = FilePath
CommandPaste.Enabled = True
Kill_or_Not = True
End Sub
Private Sub FileList_Click()
CommandCut.Enabled = True
CommandCopy.Enabled = True
If Right$(FileList.Path, 1) = "\" Then r$ = "" Else r$ = "\"
FilePath = FileList.Path + r$ + FileList.filename
FileExt = FileList.filename
Picture_Shown = True
PathLabel.Caption = "File Path : " + FilePath
NewImage.Visible = True
Sample.Picture = LoadPicture(FilePath)
NewImage.Picture = LoadPicture(FilePath)
NewImage_Width = NewImage.Width
NewImage_Height = NewImage.Height
ScaleBar.Enabled = True
ScaleBar.Value = 100
ScaleBar.Max = 500
If NewImage_Width * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Width) * 100
If NewImage_Height * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Height) * 100
Command2.Enabled = True
NewImage.Cls
Picture_Update
'W5535 H6135
End Sub
Private Sub CommandPaste_Click()
If Right$(FileList.Path, 1) = "\" Then r$ = "" Else r$ = "\"
FileCopy FilePath, FileList.Path + r$ + FileExt
CommandPaste.Enabled = False
CommandCopy.Enabled = True
CommandCut.Enabled = True
If Kill_or_Not = False Then GoTo 10
Kill FilePath: Kill_or_Not = False
FilePath = FileList.Path + r$ + FileExt
PathLabel.Caption = "File Path : " + FilePath
10 FileList.Refresh
End Sub
Private Sub Folder_Change()
FileList.Path = Folder.Path
DefauultPath = Folder.Path
CommandCut.Enabled = False
CommandCopy.Enabled = False
t = FileList.ListCount - 7
If t < 0 Then t = 0
IconScroll.Max = t
IconScroll.Value = 0
FileList.ToolTipText = Str$(FileList.ListCount) + " Items"
Update_Icons
End Sub
Private Sub Form_Load()
DefaultDrive = Drive
DefaultPath = Folder.Path
IconStart = 0
Update_Icons
End Sub
Private Sub HMove_Change()
If Dragnow = False Then MovePicture
End Sub
Private Sub HMove_Scroll()
MovePicture
End Sub
Private Sub IconList_Click(Index As Integer)
FileList.Selected(Index + IconStart) = True
End Sub
Private Sub IconScroll_Change()
IconStart = IconScroll.Value: Update_Icons
End Sub
Private Sub IconScroll_Scroll()
IconStart = IconScroll.Value: Update_Icons
End Sub
Private Sub MenuAbout_Click()
Form1.Enabled = False
frmAbout.Show
End Sub
Private Sub NewImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If ScaleBar.Enabled = False Then Exit Sub
Dragnow = True
DragX = x + HMove.Value
DragY = y + VMove.Value
End Sub
Private Sub NewImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Dragnow = False Then Exit Sub
10 H = Int(DragX - x)
V = Int(DragY - y)
If H < 0 Then H = 0 Else If H > HMove.Max Then H = HMove.Max
If V < 0 Then V = 0 Else If V > VMove.Max Then V = VMove.Max
HMove.Value = H
VMove.Value = V
MovePicture
End Sub
Private Sub NewImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dragnow = False
End Sub
Private Sub ScaleBar_Change()
Picture_Update
End Sub
Private Sub ScaleBar_Scroll()
Picture_Update
End Sub
Private Sub Picture_Update()
ScaleMeter.Caption = "Scale : " + Mid$(Str$(ScaleBar.Value), 2) + "%"
WWW = NewImage_Width * (ScaleBar.Value / 100)
If WWW > 5535 Then NewImage.Width = 5535 Else NewImage.Width = WWW
HHH = NewImage_Height * (ScaleBar.Value / 100)
If HHH > 6135 Then NewImage.Height = 6135 Else NewImage.Height = HHH
VMove.Max = HHH - NewImage.Height
HMove.Max = WWW - NewImage.Width
VMove.Value = 0
HMove.Value = 0
NewImage.PaintPicture Sample, 0, 0, WWW, HHH, 0, 0, NewImage_Width, NewImage_Height, vbSrcCopy
End Sub
Private Sub VMove_Change()
If Dragnow = False Then MovePicture
End Sub
Private Sub VMove_Scroll()
MovePicture
End Sub
Private Sub MovePicture()
s = (ScaleBar.Value / 100)
NewImage.PaintPicture Sample, 0, 0, NewImage.Width, NewImage.Height, HMove.Value / s, VMove.Value / s, NewImage.Width / s, NewImage.Height / s, vbSrcCopy
End Sub
Private Sub Update_Icons()
For i = 0 To 6
IconList(i).Visible = False
Next i
b = 0
10 b = b + 1
If b > FileList.ListCount Then Exit Sub
If b > 7 Then Exit Sub
If Right$(Folder.Path, 1) = "\" Then r$ = "" Else r$ = "\"
IconList(b - 1).Picture = LoadPicture(Folder.Path + r$ + FileList.List(b - 1 + IconStart))
IconList(b - 1).Visible = True
IconList(b - 1).ToolTipText = FileList.List(b - 1 + IconStart)
GoTo 10
End Sub
