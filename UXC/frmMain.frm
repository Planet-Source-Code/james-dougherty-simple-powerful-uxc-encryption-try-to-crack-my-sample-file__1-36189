VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UXC"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtString 
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   0
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.OptionButton optEncType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   4110
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optEncType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Top             =   4110
      Width           =   975
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecrypt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt Mode:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   4110
      Width           =   1170
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5640
      TabIndex        =   2
      Top             =   3750
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Encryptor As New UXC

Private Sub cmdDecrypt_Click()
 If optEncType(0).Value Then
  txtString.Text = Encryptor.DecryptUXC(txtString.Text, txtKey.Text)
 Else
  txtString.Text = Encryptor.ProDecryptUXC(txtString.Text, txtKey.Text)
 End If
End Sub

Private Sub cmdEncrypt_Click()
 If Len(txtKey) < 6 Or Len(txtKey) > 6 Then MsgBox "Please enter a key with 6 characters", vbOKOnly Or vbInformation, "UXC": Exit Sub
 If optEncType(0).Value Then
  txtString.Text = Encryptor.EncryptUXC(txtString.Text, txtKey.Text)
 Else
  txtString.Text = Encryptor.ProEncryptUXC(txtString.Text, txtKey.Text)
 End If
End Sub

Private Sub cmdLoad_Click()
 Dim FileNumber As Integer
 Dim tmpString As String
 
 With CD
  .Filter = "Text File (*.txt)|*.txt"
  .ShowOpen
  
  If .FileName <> vbNullString Then
   FileNumber = FreeFile()
   txtString = ""
   Open .FileName For Input As #FileNumber
    While Not EOF(FileNumber)
     Input #FileNumber, tmpString
     txtString = txtString & tmpString
    Wend
   Close
  End If
  
 End With
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Set Encryptor = Nothing
End Sub
