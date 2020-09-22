VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphics Converter"
   ClientHeight    =   3150
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Convert to &GFX"
      Height          =   450
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   2685
      Width           =   2580
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert to &BMP"
      Height          =   450
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   2685
      Width           =   2580
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   2565
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   0
      TabIndex        =   1
      Top             =   315
      Width           =   2580
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2580
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uFileName As String
Dim NewFileName As String

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case Is = 1
        Copy2GFX uFileName
        RevertBitmap NewFileName
        File1.Refresh
    Case Is = 0
        Copy2BMP uFileName
        FixBitmap NewFileName
        File1.Refresh
End Select
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    uFileName = File1.Path & "\" & File1.FileName
End Sub

Private Sub Form_Load()
    File1.Pattern = "*.BMP;*.GFX"
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Sub Copy2GFX(FileName As String)
If Right(FileName, 3) = "bmp" Or Right(FileName, 3) = "BMP" Then
    FileCopy FileName, Left(FileName, Len(FileName) - 3) & "GFX"
    NewFileName = Left(FileName, Len(FileName) - 3) & "GFX"
End If
End Sub
Sub Copy2BMP(FileName As String)

If Right(FileName, 3) = "gfx" Or Right(FileName, 3) = "GFX" Then
    FileCopy FileName, Left(FileName, Len(FileName) - 3) & "BMP"
    NewFileName = Left(FileName, Len(FileName) - 3) & "BMP"
End If
End Sub

Sub RevertBitmap(FileName As String)
    If Not FileLen(FileName) > 0 Then GoTo Error_H
    Open FileName For Binary As #1
    Put #1, 1, "."
    Put #1, 2, "."
    Put #1, 11, "."
    Put #1, 15, "."
    Close #1
Exit Sub
Error_H:
    MsgBox "File Error!", vbOKOnly, "ERROR!"
End Sub

Sub FixBitmap(FileName As String)
    If Not FileLen(FileName) > 0 Then GoTo Error_H
    Open FileName For Binary As #1
    Put #1, 1, "B"
    Put #1, 2, "M"
    Put #1, 11, "6"
    Put #1, 15, "("
    Close #1
Exit Sub
Error_H:
    MsgBox "File Error!", vbOKOnly, "ERROR!"
End Sub

