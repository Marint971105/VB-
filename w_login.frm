VERSION 5.00
Begin VB.Form w_login 
   Caption         =   "登陆_hf"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5295
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combox_id 
      Height          =   300
      ItemData        =   "w_login.frx":0000
      Left            =   2280
      List            =   "w_login.frx":000A
      TabIndex        =   7
      Text            =   "管理员"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmd_ok_txt 
      Caption         =   "登录"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txt_psw 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txt_user 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "身份"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "w_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_ok_txt_Click()
Dim strsql As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim strcon As String

'连接串
strcon = "driver={sql server}; server=(local); database=stu_hf_DB; uid=sa;pwd=sa"

con.ConnectionString = strcon '设置连接串
con.Open '建立连接

Dim n As Integer
If Combox_id = "管理员" Then
strsql = "select * from Admin_hf_user where myname='" & txt_user.Text & "'and mypsw='" & txt_psw.Text & "'"
n = 3
End If
If Combox_id = "教师" Then
strsql = "select * from Teacher_hf_user where myname='" & txt_user.Text & "'and mypsw='" & txt_psw.Text & "'"
n = 1
End If
If Combox_id = "学生" Then
strsql = "select * from Student_hf_user where myname='" & txt_user.Text & "'and mypsw='" & txt_psw.Text & "'"
n = 2
End If

'打开结果集，取得结果集记录数


rs.Open strsql, con, adOpenStatic, adLockOptimistic
Dim i As Integer
i = rs.RecordCount

If i <> 1 Then
   MsgBox "不存在的用户或密码错误！", vbCritical, "错误"
   Unload Me
Else
   If n = 2 Then
   con.Close
   w_login.Visible = False
   w_main_student.Show
   End If
   If n = 1 Then
   con.Close
   w_login.Visible = False
   w_main_teacher.Show
   End If
   If n = 3 Then
   con.Close
   Unload Me
   w_main_admin.Show
   End If

End If
End Sub

Private Sub Command2_Click()
con.Close
Unload Me

End Sub
