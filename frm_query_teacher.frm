VERSION 5.00
Begin VB.Form frm_query_teacher 
   Caption         =   "教师信息查询_hf"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6120
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt_depart 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txt_tname 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txt_tno 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txt_tsex 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmd_search 
      Caption         =   "查询"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "退出"
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "性  别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "系  别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "姓  名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "职工号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frm_query_teacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cnn As New ADODB.Connection
Dim tmpstr As String
Dim strsql As String

Private Sub cmd_exit_Click()


Unload Me


End Sub

Private Sub cmd_search_Click()
strsql = " select * from Teacher_hf where tno= '" & txt_tno.Text & "'"
rs.Open strsql, cnn
   txt_tno.Text = rs.Fields(0)
   txt_tname.Text = rs.Fields(1)
   txt_depart.Text = rs.Fields(3)
   txt_tsex.Text = rs.Fields(2)
   Me.Refresh

End Sub





Private Sub Form_load()

tmpstr = "driver={sql server}; server=(local); database=stu_hf_DB; uid=sa;pwd=sa"

cnn.ConnectionString = tmpstr
cnn.Open (tmpstr)


Me.Refresh


End Sub




