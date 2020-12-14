VERSION 5.00
Begin VB.Form frm_browse 
   Caption         =   "数据浏览_hf"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5880
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "退   出"
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmd_last 
      Caption         =   "最  后"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmd_next 
      Caption         =   "下一条"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一条"
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "置  顶"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txt_ssex 
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txt_sdept 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txt_sage 
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txt_sname 
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txt_sno 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "性   别"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "系   别"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "年   龄"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "姓   名"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "学   号"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frm_browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cnn As New ADODB.Connection
Dim strsql, tmpstr As String

Private Sub cmd_last_Click()
strsql = " select * from Student_hf order by sno desc"
rs.Open strsql, cnn
rs.MoveLast
If Not rs.EOF Then

   txt_sno.Text = rs.Fields(0)
   txt_sname.Text = rs.Fields(1)
   txt_sage.Text = rs.Fields(3)
   txt_sdept.Text = rs.Fields(4)
   txt_ssex.Text = rs.Fields(2)

   Me.Refresh

End If
End Sub

Private Sub cmd_next_Click()
strsql = " select * from Student_hf "
rs.Open strsql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly
  rs.MoveNext
If Not rs.EOF Then
   txt_sno.Text = rs.Fields(0)
   txt_sname.Text = rs.Fields(1)
   txt_sage.Text = rs.Fields(3)
   txt_sdept.Text = rs.Fields(4)
   txt_ssex.Text = rs.Fields(2)
   Me.Refresh
Else
   MsgBox "己经是最后一条记录了！"
End If
  
End Sub

Private Sub Form_load()

tmpstr = "driver={sql server}; server=(local); database=stu_hf_DB; uid=sa;pwd=sa"

cnn.ConnectionString = tmpstr
cnn.Open (tmpstr)


Me.Refresh


End Sub


