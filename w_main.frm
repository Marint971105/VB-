VERSION 5.00
Begin VB.Form w_main_student 
   Caption         =   "学生菜单_hf"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6255
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "数据检索"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "w_main_student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()
w_main.Hide
frm_insert.Show

End Sub

Private Sub Command2_Click()
w_main.Hide
frm_browse.Show

End Sub

Private Sub Command3_Click()
Unload Me
frm_query_student.Show

End Sub
