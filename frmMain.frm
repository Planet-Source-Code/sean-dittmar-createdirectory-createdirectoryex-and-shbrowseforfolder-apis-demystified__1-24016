VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CreateDirectory & CreateDirectoryEx API"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTemplate 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdTemplate 
      Caption         =   "Template Directory..."
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdCreateEx 
      Caption         =   "Create"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtCreatEx 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtCreate 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   6600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "CreateDirectoryEx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "CreateDirectory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Read the two .doc files included for a detailed
'explaination of these API calls. Thanks!

Dim SECATR As SECURITY_ATTRIBUTES
Dim folderInfo As BROWSEINFO

Private Sub cmdCreate_Click()

retval = CreateDirectory(txtCreate.Text, SECATR)
If retval <> 0 Then
    MsgBox "Sucessful"
End If

End Sub

Private Sub cmdCreateEx_Click()
If Len(txtTemplate.Text) <> 0 Then
    retval = CreateDirectoryEx(txtTemplate.Text, _
    txtCreatEx.Text, SECATR)
    If retval <> 0 Then
        MsgBox "Sucessful"
    End If
End If

End Sub

Private Sub cmdTemplate_Click()
With folderInfo
    .hwndOwner = Me.hWnd
    .pszDisplayName = Space(260)
    retval = SHBrowseForFolder(folderInfo)
    txtTemplate.Text = "c:\" & .pszDisplayName & "\"
End With

End Sub

Private Sub Form_Load()
txtCreate.Text = App.Path
txtCreatEx.Text = App.Path
End Sub
