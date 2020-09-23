VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Browse For Folder Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowseForFolder 
      Caption         =   "Display Browse For Folder"
      Height          =   705
      Left            =   900
      TabIndex        =   0
      Top             =   1140
      Width           =   2235
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Note: The BrowseForFolder can show also files if the Flag BIF_BROWSEINCLUDEFILES is selected"
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   4005
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   585
      Left            =   60
      TabIndex        =   1
      Top             =   1890
      Width           =   4005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BrowseForFolder by Max Raskin, February 25 2000

Private Sub cmdBrowseForFolder_Click()
    Dim ReturnValue As String
    ReturnValue = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)
    If ReturnValue <> "" Then
      lblInfo.Caption = "Path Selected: " & ReturnValue
    Else
      lblInfo.Caption = "Cancel selected or the folder/file type selected isn't from the file system"
    End If
End Sub
