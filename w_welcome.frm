VERSION 5.00
Begin VB.Form w_welcome 
   Caption         =   "��ӭ_hf"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   7470
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��¼"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "��ӭ����ѧ����Ϣ����ϵͳ"
      BeginProperty Font 
         Name            =   "��������"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
End
Attribute VB_Name = "w_welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
w_login.Show
w_welcome.Hide

End Sub

