VERSION 5.00
Begin VB.Form w_main_admin 
   Caption         =   "����Ա�˵�_hf"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5130
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "��ʦ�������"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ѧ������¼��"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ѧ���������"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʦ����¼��"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "w_main_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

frm_insert.Show

End Sub

Private Sub Command2_Click()
Unload Me
frm_browse.Show
End Sub

Private Sub Command3_Click()
Unload Me

frm_insert.Show
End Sub

Private Sub Command5_Click()
Unload Me
frm_browse.Show
End Sub
