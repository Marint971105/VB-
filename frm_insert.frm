VERSION 5.00
Begin VB.Form frm_insert 
   Caption         =   "ѧ������¼��_hf"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7770
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmd_exit 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox Combo_ssex 
      Height          =   300
      ItemData        =   "frm_insert.frx":0000
      Left            =   3360
      List            =   "frm_insert.frx":0007
      TabIndex        =   9
      Text            =   "��"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txt_sdept 
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txt_sage 
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txt_sname 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txt_sno 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ϵ  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ѧ  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frm_insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strcon As String
Dim strsql As String

Private Sub cmd_exit_Click()
 con.Close
 Unload Me
 w_main.Show
 
End Sub

Private Sub cmd_ok_Click()
Dim strsql As String
Dim i As Integer

strsql = " insert into Student_hf values ('" + txt_sno.Text + "','" + txt_sname.Text + "','" + Combo_ssex.Text + "'," + txt_sage.Text + ",'" + txt_sdept.Text + "')"
rs.Open strsql, con
i = con.State

 If i = 1 Then
         MsgBox "�����ѳɹ����"
 End If

End Sub

Private Sub lable1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_load()
Dim tmpstr As String

tmpstr = "driver={sql server}; server=(local); database=stu_hf_DB; uid=sa;pwd=sa"

con.ConnectionString = tmpstr
con.Open


Me.Refresh


End Sub
