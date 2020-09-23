VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Encryption"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "By Sergio del Rio"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        Dim encrypt
            For e = 1 To 5
                encrypt = Encode(Text1.Text)
            Text1.Text = encrypt
            Next e
        
End Sub

Private Sub Command2_Click()
        Dim decrypt
            For d = 1 To 5
                decrypt = DeCode(Text1.Text)
                Text1.Text = decrypt
            Next d
End Sub
