VERSION 5.00
Begin VB.Form w_main_admin 
   Caption         =   "管理员菜单_hf"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5130
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "教师数据浏览"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "学生数据录入"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "学生数据浏览"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "教师数据录入"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
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
