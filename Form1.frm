VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "诉讼费计算器"
   ClientHeight    =   4632
   ClientLeft      =   5508
   ClientTop       =   3228
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4632
   ScaleWidth      =   6000
   Begin VB.CheckBox Check2 
      Caption         =   "调解/撤诉"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   4032
      TabIndex        =   14
      Top             =   3456
      Width           =   1356
   End
   Begin VB.CheckBox Check2 
      Caption         =   "反诉/有独三"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   4032
      TabIndex        =   13
      Top             =   3072
      Width           =   1485
   End
   Begin VB.CheckBox Check2 
      Caption         =   "简易程序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   4032
      TabIndex        =   12
      Top             =   2730
      Width           =   1170
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   330
      Left            =   315
      TabIndex        =   11
      Top             =   2835
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2940
      TabIndex        =   8
      Text            =   "0"
      Top             =   2310
      Width           =   2220
   End
   Begin VB.CommandButton Command3 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2205
      TabIndex        =   7
      Top             =   3780
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2940
      TabIndex        =   5
      Text            =   "0"
      Top             =   1365
      Width           =   2220
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":0000
      Left            =   315
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   2310
      Width           =   2220
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IntegralHeight  =   0   'False
      ItemData        =   "Form1.frx":0004
      Left            =   315
      List            =   "Form1.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1365
      Width           =   1590
   End
   Begin VB.Label Label7 
      Caption         =   "其他选项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2940
      TabIndex        =   15
      Top             =   2832
      Width           =   948
   End
   Begin VB.Label Label6 
      Caption         =   "元"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5355
      TabIndex        =   10
      Top             =   2310
      Width           =   225
   End
   Begin VB.Label Label5 
      Caption         =   "费用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2940
      TabIndex        =   9
      Top             =   1890
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "元"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5250
      TabIndex        =   6
      Top             =   1470
      Width           =   225
   End
   Begin VB.Label Label3 
      Caption         =   "标的额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2940
      TabIndex        =   4
      Top             =   945
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "案件类型"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   2
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "费用类型"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   1
      Top             =   945
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
k = 1000
w = 10000
p = 0.01
MsgBox "依据《诉讼费用交纳办法》编制，实际费用标准请以当地法院为准", vbOKOnly, "提示"
Form2.Show 1
End Sub

Private Sub Combo1_Click()
Combo2.Clear
If Combo1.ListIndex = 0 Then
    Combo2.AddItem "财产案件"
    Combo2.AddItem "离婚案件"
    Combo2.AddItem "侵犯人格权案件"
    Combo2.AddItem "知识产权民事案件"
    Combo2.AddItem "劳动争议案件"
    Combo2.AddItem "其他非财产案件"
    Combo2.AddItem "管辖权异议"
    Combo2.AddItem "商标、专利、海事行政案件"
    Combo2.AddItem "其他行政案件"
ElseIf Combo1.ListIndex = 1 Then
    Combo2.AddItem "申请执行"
    Combo2.AddItem "申请财产保全"
    Combo2.AddItem "申请支付令"
    Combo2.AddItem "申请公示催告"
    Combo2.AddItem "申请撤销仲裁或认定仲裁协议"
    Combo2.AddItem "破产案件"
    Combo2.AddItem "申请设立海事赔偿责任限制基金"
    Combo2.AddItem "申请海事强制令"
    Combo2.AddItem "申请申请船舶优先权催告"
    Combo2.AddItem "申请海事债权登记"
    Combo2.AddItem "申请共同海损理算"
End If
End Sub

Private Sub Combo2_Click()
If Combo1.ListIndex = 0 Then
    If Combo2.ListIndex = 1 Then
        Check1.Visible = True
        Check1.Caption = "涉及财产分割"
    ElseIf Combo2.ListIndex = 2 Then
        Check1.Visible = True
        Check1.Caption = "涉及损害赔偿"
    ElseIf Combo2.ListIndex = 3 Then
        Check1.Visible = True
        Check1.Caption = "涉及争议金额"
    Else
        Check1.Visible = False
        Check1.value = 0
    End If
ElseIf Combo1.ListIndex = 1 Then
    If Combo2.ListIndex = 0 Then
        Check1.Visible = True
        Check1.Caption = "涉及执行金额"
    Else
        Check1.Visible = False
        Check1.value = 0
    End If
End If
End Sub

Private Sub Command3_Click()
If IsNumeric(Text1.Text) = False Then           '数值合法性检验
    MsgBox "请输入合法的数值", vbOKOnly, "提示"
    Text1.Text = 0
    Exit Sub
ElseIf Text1.Text < 0 Then
    MsgBox "请输入合法的数值", vbOKOnly, "提示"
    Text1.Text = 0
    Exit Sub
End If

half = False
For i = 0 To 2
    If Check2(i).value = 1 Then
        half = True
        Exit For
    End If
Next i
If Combo1.ListIndex = 0 Then
    Select Case Combo2.ListIndex
    Case 0
        Text2.Text = sl0(Val(Text1.Text))
    Case 1
        Text2.Text = sl1(Val(Text1.Text))
    Case 2
        Text2.Text = sl2(Val(Text1.Text))
    Case 3
        Text2.Text = sl3(Val(Text1.Text))
    Case 4
        sl4 (0)
    Case 5
        sl5 (0)
    Case 6
        sl6 (0)
    Case 7
        sl7 (0)
    Case 8
        sl8 (0)
    Case Else
        MsgBox "请选择一个项目", vbOKOnly, "提示"
        Exit Sub
    End Select
ElseIf Combo1.ListIndex = 1 Then
    Select Case Combo2.ListIndex
    Case 0
        Text2.Text = sq0(Val(Text1.Text))
    Case 1
        Text2.Text = sq1(Val(Text1.Text))
    Case 2
        Text2.Text = sq2(Val(Text1.Text))
    Case 3
        sq3 (0)
    Case 4
        sq4 (0)
    Case 5
        Text2.Text = sq5(Val(Text1.Text))
    Case 6
        Text2.Text = sq6(Val(Text1.Text))
    Case 7
        Text2.Text = sq7(Val(Text1.Text))
    Case 8
        Text2.Text = sq8(Val(Text1.Text))
    Case 9
        sq9 (0)
    Case 10
        sq10 (0)
    Case Else
        MsgBox "请选择一个项目", vbOKOnly, "提示"
        Exit Sub
    End Select
End If

For i = 0 To 2
    If Check2(i).value = 1 Then
        Text2.Text = Text2.Text / 2
        Exit For
    End If
Next i
End Sub

Function sl0(ByVal value As Double)
If value < 1 * w Then
    sl0 = 50
ElseIf value < 10 * w Then
    sl0 = 50 + (value - 1 * w) * 2.5 * p
ElseIf value < 20 * w Then
    sl0 = 2300 + (value - 10 * w) * 2 * p
ElseIf value < 50 * w Then
    sl0 = 4300 + (value - 20 * w) * 1.5 * p
ElseIf value < 100 * w Then
    sl0 = 8800 + (value - 50 * w) * 1 * p
ElseIf value < 200 * w Then
    sl0 = 13800 + (value - 100 * w) * 0.9 * p
ElseIf value < 500 * w Then
    sl0 = 22800 + (value - 200 * w) * 0.8 * p
ElseIf value < 1000 * w Then
    sl0 = 46800 + (value - 500 * w) * 0.7 * p
ElseIf value < 2000 * w Then
    sl0 = 81800 + (value - 1000 * w) * 0.6 * p
Else
    sl0 = 141800 + (value - 2000 * w) * 0.5 * p
End If
End Function

Function sl1(ByVal value As Double)
If Check1.value = 0 Then
    sl1 = "50 - 300"
Else
    If value < 20 * w Then
        sl1 = "50 - 300"
    Else
        sl1 = "50 - 300 & " & (value - 20 * w) * 0.5 * p
    End If
End If
End Function

Function sl2(ByVal value As Double)
If Check1.value = 0 Then
    sl2 = "100 - 500"
Else
    If value < 5 * w Then
        sl2 = "100 - 500"
    Else
        If value < 10 * w Then
            sl2 = "100 - 500 & " & (value - 5 * w) * 1 * p
        Else
            sl2 = "100 - 500 & " & 500 + (value - 10 * w) * 0.5 * p
        End If
    End If
End If
End Function

Function sl3(ByVal value As Double)
If Check1.value = 0 Then
    sl3 = "500 -1000"
Else
    sl3 = sl0(value)
End If
End Function

Function sl4(ByVal value As Double)
If half = True Then
    MsgBox "每件5元", vbOKOnly, "提示"
Else
    MsgBox "每件10元", vbOKOnly, "提示"
End If
End Function

Function sl5(ByVal value As Double)
If half = True Then
    MsgBox "每件25 - 50元", vbOKOnly, "提示"
Else
    MsgBox "每件50 - 100元", vbOKOnly, "提示"
End If
End Function

Function sl6(ByVal value As Double)
MsgBox "异议不成立时，" & vbCrLf & "每件50 - 100元", vbOKOnly, "提示"
End Function

Function sl7(ByVal value As Double)
If half = True Then
    MsgBox "每件50元", vbOKOnly, "提示"
Else
    MsgBox "每件100元", vbOKOnly, "提示"
End If
End Function

Function sl8(ByVal value As Double)
If half = True Then
    MsgBox "每件25元", vbOKOnly, "提示"
Else
    MsgBox "每件50元", vbOKOnly, "提示"
End If
End Function

Function sq0(ByVal value As Double)
If Check1.value = 0 Then
    sq0 = "50 - 500"
Else
    If value < 1 * w Then
        sq0 = 50
    ElseIf value < 50 * w Then
        sq0 = 50 + (value - 1 * w) * 1.5 * p
    ElseIf value < 500 * w Then
        sq0 = 7400 + (value - 50 * w) * 1 * p
    ElseIf value < 1000 * w Then
        sq0 = 52400 + (value - 500 * w) * 0.5 * p
    Else
        sq0 = 77400 + (value - 1000 * w) * 0.1 * p
    End If
End If
End Function

Function sq1(ByVal value As Double)
If value < 1 * k Then
    sq1 = 30
ElseIf value < 10 * w Then
    sq1 = 30 + (value - 1 * k) * 1 * p
Else
    sq1 = 525 + (value - 10 * w) * 0.5 * p
End If
sq1 = IIf(sq1 > 5000, 5000, sq1)
End Function

Function sq2(ByVal value As Double)
sq2 = sl0(value) / 3
End Function

Function sq3(ByVal value As Double)
If half = True Then
    MsgBox "每件50元", vbOKOnly, "提示"
Else
    MsgBox "每件100元", vbOKOnly, "提示"
End If
End Function

Function sq4(ByVal value As Double)
If half = True Then
    MsgBox "每件200元", vbOKOnly, "提示"
Else
    MsgBox "每件400元", vbOKOnly, "提示"
End If
End Function

Function sq5(ByVal value As Double)
sq5 = sl0(value) / 2
sq5 = IIf(sq5 > 30 * w, 30 * w, sq5)
End Function

Function sq6(ByVal value As Double)
sq6 = "1000 - 10000"
End Function

Function sq7(ByVal value As Double)
sq7 = "1000 - 5000"
End Function

Function sq8(ByVal value As Double)
sq8 = "1000 - 5000"
End Function

Function sq9(ByVal value As Double)
If half = True Then
    MsgBox "每件500元", vbOKOnly, "提示"
Else
    MsgBox "每件1000元", vbOKOnly, "提示"
End If
End Function

Function sq10(ByVal value As Double)
If half = True Then
    MsgBox "每件500元", vbOKOnly, "提示"
Else
    MsgBox "每件1000元", vbOKOnly, "提示"
End If
End Function
