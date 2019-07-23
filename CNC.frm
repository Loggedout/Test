VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "数控插补仿真软件"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox lqy 
      Height          =   350
      Left            =   2280
      TabIndex        =   21
      Top             =   9600
      Width           =   1000
   End
   Begin VB.TextBox Motion 
      Height          =   350
      Left            =   2280
      TabIndex        =   20
      Top             =   9000
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清除屏幕"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2040
      TabIndex        =   19
      Top             =   6840
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   120
   End
   Begin VB.TextBox Y0 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2280
      TabIndex        =   18
      Text            =   "5000"
      Top             =   8400
      Width           =   1000
   End
   Begin VB.TextBox X0 
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
      Left            =   2280
      TabIndex        =   17
      Text            =   "8000"
      Top             =   7800
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   9525
      Left            =   4440
      ScaleHeight     =   9465
      ScaleMode       =   0  'User
      ScaleWidth      =   12148.41
      TabIndex        =   13
      Top             =   360
      Width           =   12645
   End
   Begin VB.TextBox LuJin 
      Height          =   350
      Left            =   0
      TabIndex        =   12
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox NCtext 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CommandButton MoNi 
      Caption         =   "路径模拟"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   0
      TabIndex        =   10
      Top             =   6840
      Width           =   1755
   End
   Begin VB.CommandButton DuQu 
      Caption         =   "读取文件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   2000
   End
   Begin VB.TextBox YAnytime 
      Height          =   350
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   1500
   End
   Begin VB.TextBox XAnytime 
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Width           =   1500
   End
   Begin VB.TextBox Jinji 
      Height          =   350
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   1500
   End
   Begin VB.TextBox ZhuZhou 
      Height          =   350
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "冷却液开关"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   23
      Top             =   9600
      Width           =   1995
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "主轴状态"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   22
      Top             =   9000
      Width           =   1995
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Y0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   16
      Top             =   8400
      Width           =   1995
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "X0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   7800
      Width           =   1995
   End
   Begin VB.Label Label6 
      Caption         =   "工件坐标系"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Top             =   7320
      Width           =   1995
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "实时坐标"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "进给速度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "主轴转速"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrFile As String
Dim intFile As Integer
Dim strData As String
Dim NcLine() As String
Dim length As Integer
Dim InitX As Double '起始点X'
Dim InitY As Double '起始点Y'
Dim NextX As Double '终止点X'
Dim NextY As Double '终止点Y'
Dim InitS As Double '主轴转速'
Dim InitF As Double '进给速度'
Dim InitG901 As String '相对或绝或对坐标'
Dim InitG0123 As String 'G00或G01或G02或G03指令'
Dim InitI As Double 'X方向相对圆心位置'
Dim InitJ As Double 'Y方向相对圆心位置'
Dim Time As Integer
Dim judge As Integer



'Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)'

'该函数用于获取主轴速度InitS'
Function S(EachOn As String)
    S = Val(Mid(EachOn, 2))
    InitS = S
    ZhuZhou.Text = InitS
End Function

'该函数用于获取进给速度InitF'
Function F(EachOn As String)
    F = Val(Mid(EachOn, 2))
    InitF = F
    Jinji.Text = InitF
End Function

'该函数用于获取或计算下一点在X方向绝对坐标NextX'
Function X(EachOn As String)
    X = Val(Mid(EachOn, 2))
    If InitG901 = "G90" Then  'G90下获取绝对坐标'
        NextX = X
    Else
        NextX = X + InitX 'G91下获取相对坐标'
    End If
End Function

'该函数用于获取或计算下一点在Y方向绝对坐标NextY'
Function Y(EachOn As String)
    Y = Val(Mid(EachOn, 2))
    If InitG901 = "G90" Then 'G90下获取绝对坐标'
        NextY = Y
    Else
        NextY = Y + InitY 'G91下获取相对坐标'
    End If
End Function

'该函数用于获取圆心在X方向相对坐标'
Function I(EachOn As String)
    InitI = Val(Mid(EachOn, 2))
End Function

'该函数用于获取圆心在Y方向相对坐标'
Function J(EachOn As String)
    InitJ = Val(Mid(EachOn, 2))
End Function
Function G(EachOn As String)
    If EachOn = "G90" Then
        InitG901 = "G90"
    ElseIf EachOn = "G91" Then
        InitG901 = "G91"
    Else
        InitG0123 = EachOn
    End If
End Function

'该函数用于获取主轴和冷却液状态'
'Function M(EachOn As String)
'    If Val(Mid(EachOn, 3)) = 3 Then
'        Motion.Text = "正转"
'    ElseIf Val(Mid(EachOn, 3)) = 4 Then
'        Motion.Text = "反转"
'    ElseIf Val(Mid(EachOn, 3)) = 5 Then
'        Motion.Text = "停转"
'    ElseIf Val(Mid(EachOn, 3)) = 7 Then
'        lqy.Text = "打开"
'    ElseIf Val(Mid(EachOn, 3)) = 8 Then
'        lqy.Text = "打开"
'    Else
'        lqy.Text = "关闭"
'    End If
'End Function

Function M(EachOn As String)
    If Val(Mid(EachOn, 3)) = 3 Then         'M03指令
        Motion.Text = "正转"
    ElseIf Val(Mid(EachOn, 3)) = 4 Then     'M04指令
        Motion.Text = "反转"
    ElseIf Val(Mid(EachOn, 3)) = 5 Then     'M05指令
        Motion.Text = "停转"
    ElseIf Val(Mid(EachOn, 3)) = 7 Then     'M07指令
        lqy.Text = "打开"
    ElseIf Val(Mid(EachOn, 3)) = 8 Then     'M08指令
        lqy.Text = "打开"
    ElseIf Val(Mid(EachOn, 3)) = 9 Then     'M09指令
        lqy.Text = "打开"
    ElseIf Val(Mid(EachOn, 3)) = 2 Then     'M02指令
        lqy.Text = "关闭"
        Motion.Text = "停转"
        Call G00(InitX, InitY, 0, 0)
        judge = 1
    Else                                    'M30指令
        lqy.Text = "关闭"
        Motion.Text = "停转"
        Call G00(InitX, InitY, X0, Y0)
        judge = 2
    End If
End Function

'该函数为快速进给指令'
Function G00(IX As Double, IY As Double, NX As Double, NY As Double)
    Timer1.Enabled = True
    Timer1.Interval = 10
    Dim deltaX As Double
    Dim deltaY As Double
    Dim N As Integer
    N = 50
    deltaX = (NX - IX) / N
    deltaY = (NY - IY) / N
    Dim Count As Double
    Count = 0
    Do While Count <= N - 1
        DoEvents
        If Time Then
            Picture1.Line (IX + deltaX * Count, IY + deltaY * Count)-(IX + deltaX * (Count + 1), IY + deltaY * (Count + 1)), vbRed
            XAnytime.Text = IX + deltaX * (Count + 1)
            YAnytime.Text = IY + deltaY * (Count + 1)
            Time = 0
            Count = Count + 1
        End If
    Loop
    InitX = IX + deltaX * N
    InitY = IY + deltaY * N
End Function

'该函数伟直线插补指令'
Function G01(IX As Double, IY As Double, NX As Double, NY As Double)
    Timer1.Enabled = True
    Timer1.Interval = 10
    Dim MaiCh As Integer
    Dim CounX As Integer
    Dim CounY As Integer
    Dim Ye As Double
    Dim Xe As Double
    Xe = NX - IX
    Ye = NY - IY
    MaiCh = 10
    CounX = Int(Abs(NX - IX) / MaiCh)
    CounY = Int(Abs(NY - IY) / MaiCh)
    Dim ZFX As Integer
    Dim ZFY As Integer
    If NX > IX Then
        ZFX = 1
    Else
        ZFX = -1
    End If
    If NY > IY Then
        ZFY = 1
    Else
        ZFY = -1
    End If
    Dim Fm As Double
    Fm = 0
    Dim Count As Double
    Count = CounX + CounY
    Do While Count > 0
        DoEvents
        If Time Then
            If Fm >= 0 Then
                Picture1.Line (IX, IY)-(IX + ZFX * MaiCh, IY)
                IX = IX + ZFX * MaiCh
                XAnytime.Text = IX
                Fm = Fm - Abs(Ye)
                
            Else
                Picture1.Line (IX, IY)-(IX, IY + ZFY * MaiCh)
                IY = IY + ZFY * MaiCh
                YAnytime.Text = IY
                Fm = Fm + Abs(Xe)
            End If
            Time = 0
            Count = Count - 1
        End If
    Loop
    Timer1.Enabled = False
    InitX = IX
    InitY = IY
End Function


'该函数为逆圆插补指令'
Function G03(IX As Double, IY As Double, NX As Double, NY As Double, II As Double, JJ As Double)
    
    Dim deltaX As Double
    Dim deltaY As Double
    Dim dd As Double
    
    dd = InitF / 60                           '定义脉冲当量
                                        
    Dim Count As Integer
    
    Dim PX As Double                          '定义起点X相对于圆心点的PX
    Dim PY As Double                          '定义起点Y相对于圆心点的PY
    Dim PNX As Double                         '定义终点X相对于圆心点的PX
    Dim PNY As Double                         '定义终点Y相对于圆心点的PY
    
    Dim ZJDX As Double                         '定义中间过渡点X绝对坐标
    Dim ZJDY As Double                         '定义中间过渡点Y绝对坐标
    
    Dim RR2 As Double                           '定义圆弧半径的平方
   
    
    II = II + IX                               '将圆心I的相对坐标变为绝对坐标
    JJ = JJ + IY                               '将圆心J的相对坐标变为绝对坐标
    RR2 = Sqr((IX - II) ^ 2 + (IY - JJ) ^ 2)         '计算圆弧半径
    
    PX = IX - II                              '起始点相对于圆点的X坐标
    PY = IY - JJ                              '起始点相对于圆点的Y坐标
    PNX = NX - II                             '终止点相对于圆点的X坐标
    PNY = NY - JJ                             '终止点相对于圆点的Y坐标
            
    Dim XiangXian As Integer                  '用来存储起始点的象限
            
             '起始点在第一象限的情况'
            If PX > 0 And PY >= 0 And PNX >= 0 And PNY > 0 Then                    '起始点在第一象限的情况，终止点在第一象限'
                XiangXian = 11
 
            ElseIf PX > 0 And PY >= 0 And PNX < 0 And PNY >= 0 Then                '起始点在第一象限的情况，终止点在第二象限'
                XiangXian = 12
     
            ElseIf PX > 0 And PY >= 0 And PNX <= 0 And PNY < 0 Then                '起始点在第一象限的情况，终止点在第三象限'
                XiangXian = 13
 
            ElseIf PX > 0 And PY >= 0 And PNX > 0 And PNY <= 0 Then                '起始点在第一象限的情况，终止点在第四象限'
                XiangXian = 14
                
                
             '起始点在第二象限的情况'
            ElseIf PX <= 0 And PY > 0 And PNX < 0 And PNY >= 0 Then               '起始点在第二象限的情况，终止点在第二象限'
                XiangXian = 21
            
            ElseIf PX <= 0 And PY > 0 And PNX <= 0 And PNY < 0 Then                '起始点在第二象限的情况，终止点在第三象限'
                XiangXian = 22
                
            ElseIf PX <= 0 And PY > 0 And PNX > 0 And PNY <= 0 Then                 '起始点在第二象限的情况，终止点在第四象限'
                XiangXian = 23
                
            ElseIf PX <= 0 And PY > 0 And PNX >= 0 And PNY > 0 Then                '起始点在第二象限的情况，终止点在第一象限'
                XiangXian = 24
            
            
            '起始点在第三象限的情况'
            ElseIf PX < 0 And PY <= 0 And PNX <= 0 And PNY < 0 Then                '起始点在第三象限的情况，终止点在第三象限'
                XiangXian = 31
             
            ElseIf PX < 0 And PY <= 0 And PNX > 0 And PNY <= 0 Then                '起始点在第三象限的情况，终止点在第四象限'
                XiangXian = 32
                
            ElseIf PX < 0 And PY <= 0 And PNX >= 0 And PNY > 0 Then                 '起始点在第三象限的情况，终止点在第一象限'
                XiangXian = 33
                
            ElseIf PX < 0 And PY <= 0 And PNX < 0 And PNY >= 0 Then                '起始点在第三象限的情况，终止点在第二象限'
                XiangXian = 34
               
               
            '起始点在第四象限的情况'
            ElseIf PX >= 0 And PY < 0 And PNX > 0 And PNY <= 0 Then                '起始点在第四象限的情况，终止点在第四象限'
                XiangXian = 41
                
            ElseIf PX >= 0 And PY < 0 And PNX >= 0 And PNY > 0 Then                '起始点在第四象限的情况，终止点在第一象限'
                XiangXian = 42
                
            ElseIf PX >= 0 And PY < 0 And PNX < 0 And PNY >= 0 Then                 '起始点在第四象限的情况，终止点在第二象限'
                XiangXian = 43
                
            ElseIf PX >= 0 And PY < 0 And PNX <= 0 And PNY < 0 Then                '起始点在第四象限的情况，终止点在第三象限'
                XiangXian = 44
 
            End If
            
            '根据起始点和终止点相对于圆心的所在的象限选择插补方案'
                   Select Case XiangXian
                        '第一大类'
                        Case "11"
                            If IX <= NX Then
                                '第一段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan2 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                                '第五段圆弧'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                Else
                                
                                '第二情况'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                End If
                                
                        Case "12"
                            '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                        Case "13"
                            '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                        Case "14"
                        '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                        
                        
                        '第二大类'
                        Case "21"
                            If IX <= NX Then
                                '第一段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            
                            '第五段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            Else
                            
                            '第二情况'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "22"
                            '第一段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                        Case "23"
                            '第一段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                        Case "24"
                            '第一段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            
                        '第三大类'
                        Case "31"
                            If IX >= NX Then
                                '第一段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan2 IX, IY, II, JJ, dd, Count
                                '第五段圆弧'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                            Else
                            
                                '第二情况'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "32"
                            '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                        Case "33"
                            '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                        Case "34"
                        '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            
                        
                        '第四大类'
                        Case "41"
                            If IX >= NX Then
                                '第一段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan2 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                                '第五段圆弧'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                            Else
                            '第二情况'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "42"
                            '第一段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                        Case "43"
                            '第一段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                        Case "44"
                        '第一段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                    End Select
    
    InitX = NextX
    InitY = NextY
    InitI = 0                '这样可以缺省I0
    InitJ = 0                '这样可以缺省J0
End Function
'G03第一个象限的方案'
            Function G03fangan1(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX - ddd, IYY), vbRed
                                IXX = IXX - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY + ddd), vbRed
                                IYY = IYY + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function


'G03第二个象限的方案'
            Function G03fangan2(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX, IYY - ddd), vbRed
                                IYY = IYY - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX - ddd, IYY), vbRed
                                IXX = IXX - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function
            

'G03第三个象限的方案'
            Function G03fangan3(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX + ddd, IYY), vbRed
                                IXX = IXX + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY - ddd), vbRed
                                IYY = IYY - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function


'G03第四个象限的方案'
            Function G03fangan4(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX, IYY + ddd), vbRed
                                IYY = IYY + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX + ddd, IYY), vbRed
                                IXX = IXX + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function
 
 
 '该函数为顺圆插补指令'
Function G02(IX As Double, IY As Double, NX As Double, NY As Double, II As Double, JJ As Double)
    
    Dim deltaX As Double
    Dim deltaY As Double
    Dim dd As Double
    
    dd = InitF / 60                           '定义脉冲当量
                                        
    Dim Count As Integer
    
    Dim PX As Double                          '定义起点X相对于圆心点的PX
    Dim PY As Double                          '定义起点Y相对于圆心点的PY
    Dim PNX As Double                         '定义终点X相对于圆心点的PX
    Dim PNY As Double                         '定义终点Y相对于圆心点的PY
    
    Dim ZJDX As Double                         '定义中间过渡点X绝对坐标
    Dim ZJDY As Double                         '定义中间过渡点Y绝对坐标
    
    Dim RR2 As Double                           '定义圆弧半径的平方
   
    
    II = II + IX                               '将圆心I的相对坐标变为绝对坐标
    JJ = JJ + IY                               '将圆心J的相对坐标变为绝对坐标
    RR2 = Sqr((IX - II) ^ 2 + (IY - JJ) ^ 2)         '计算圆弧半径
    
    PX = IX - II                              '起始点相对于圆点的X坐标
    PY = IY - JJ                              '起始点相对于圆点的Y坐标
    PNX = NX - II                             '终止点相对于圆点的X坐标
    PNY = NY - JJ                             '终止点相对于圆点的Y坐标
            
    Dim XiangXian As Integer                  '用来存储起始点的象限
            
             '起始点在第一象限的情况'
            If PX >= 0 And PY > 0 And PNX >= 0 And PNY > 0 Then                   '起始点在第一象限的情况，终止点在第一象限'
                XiangXian = 11
 
            ElseIf PX >= 0 And PY > 0 And PNX < 0 And PNY >= 0 Then                '起始点在第一象限的情况，终止点在第二象限'
                XiangXian = 12
     
            ElseIf PX >= 0 And PY > 0 And PNX <= 0 And PNY < 0 Then                '起始点在第一象限的情况，终止点在第三象限'
                XiangXian = 13
 
            ElseIf PX >= 0 And PY > 0 And PNX > 0 And PNY <= 0 Then                '起始点在第一象限的情况，终止点在第四象限'
                XiangXian = 14
                
                
             '起始点在第二象限的情况'
            ElseIf PX <= 0 And PY >= 0 And PNX < 0 And PNY >= 0 Then              '起始点在第二象限的情况，终止点在第二象限'
                XiangXian = 21
            
            ElseIf PX < 0 And PY >= 0 And PNX <= 0 And PNY < 0 Then                '起始点在第二象限的情况，终止点在第三象限'
                XiangXian = 22
                
            ElseIf PX < 0 And PY >= 0 And PNX > 0 And PNY <= 0 Then                 '起始点在第二象限的情况，终止点在第四象限'
                XiangXian = 23
                
            ElseIf PX < 0 And PY >= 0 And PNX >= 0 And PNY > 0 Then                '起始点在第二象限的情况，终止点在第一象限'
                XiangXian = 24
            
            
            '起始点在第三象限的情况'
            ElseIf PX <= 0 And PY < 0 And PNX <= 0 And PNY < 0 Then               '起始点在第三象限的情况，终止点在第三象限'
                XiangXian = 31
             
            ElseIf PX <= 0 And PY < 0 And PNX > 0 And PNY <= 0 Then                '起始点在第三象限的情况，终止点在第四象限'
                XiangXian = 32
                
            ElseIf PX <= 0 And PY < 0 And PNX >= 0 And PNY > 0 Then                 '起始点在第三象限的情况，终止点在第一象限'
                XiangXian = 33
                
            ElseIf PX <= 0 And PY < 0 And PNX < 0 And PNY >= 0 Then                '起始点在第三象限的情况，终止点在第二象限'
                XiangXian = 34
               
               
            '起始点在第四象限的情况'
            ElseIf PX > 0 And PY <= 0 And PNX > 0 And PNY <= 0 Then                '起始点在第四象限的情况，终止点在第四象限'
                XiangXian = 41
                
            ElseIf PX > 0 And PY <= 0 And PNX >= 0 And PNY > 0 Then                '起始点在第四象限的情况，终止点在第一象限'
                XiangXian = 42
                
            ElseIf PX > 0 And PY <= 0 And PNX < 0 And PNY >= 0 Then                 '起始点在第四象限的情况，终止点在第二象限'
                XiangXian = 43
                
            ElseIf PX > 0 And PY <= 0 And PNX <= 0 And PNY < 0 Then                '起始点在第四象限的情况，终止点在第三象限'
                XiangXian = 44
 
            End If
            
            '根据起始点和终止点相对于圆心的所在的象限选择插补方案'
                   Select Case XiangXian
                        '第一大类'
                        Case "11"                                                 '起始点在第一象限的情况，终止点在第一象限'
                            If IX >= NX Then
                                '第一段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '第五段圆弧'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                            
                                Else
                                
                                '第二情况'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                            
                                End If
                                 
                        Case "12"                                                             '起始点在第一象限的情况，终止点在第二象限'
                                '第一段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                             '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                        Case "13"                                                             '起始点在第一象限的情况，终止点在第三象限'
                            '第一段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                        Case "14"                                                                '起始点在第一象限的情况，终止点在第四象限'
                        '第一段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                        
                        
                        '第二大类'
                        Case "21"                                                 '起始点在第二象限的情况，终止点在第二象限'
                            If IX >= NX Then
                                '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                            
                            '第五段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            Else
                            
                            '第二情况'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "22"                                          '起始点在第二象限的情况，终止点在第三象限'
                            '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                            '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                
                        Case "23"                                                  '起始点在第二象限的情况，终止点在第四象限'
                            '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                        Case "24"                                           '起始点在第二象限的情况，终止点在第一象限'
                            '第一段圆弧'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            
                        '第三大类'
                        Case "31"                                            '起始点在第三象限的情况，终止点在第三象限'
                            If IX <= NX Then
                                '第一段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '第五段圆弧'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                            Else
                            
                                '第二情况'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "32"                                                     '起始点在第三象限的情况，终止点在第四象限'
                            '第一段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                        Case "33"                                                    '起始点在第三象限的情况，终止点在第一象限'
                            '第一段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                        Case "34"                                                '起始点在第三象限的情况，终止点在第二象限'
                        '第一段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                            '第二段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            
                        
                        '第四大类'
                        Case "41"                                                         '起始点在第四象限的情况，终止点在第四象限'
                            If IX <= NX Then
                                '第一段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '第五段圆弧'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                            Else
                            '第二情况'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "42"                                                            '起始点在第四象限的情况，终止点在第一象限'
                            '第一段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '第四段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                        Case "43"                                                          '起始点在第四象限的情况，终止点在第二象限'
                           '第一段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '第三段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                        Case "44"                                                    '起始点在第四象限的情况，终止点在第三象限'
                        '第一段圆弧'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '第二段圆弧'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                    End Select
    
    InitX = NextX
    InitY = NextY
    InitI = 0                '缺省I0
    InitJ = 0                '缺省J0
End Function
'G02第一个象限的方案'
            Function G02XX1(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX, IYY - ddd), vbRed
                                IXX = IXX
                                IYY = IYY - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX + ddd, IYY), vbRed
                                IXX = IXX + ddd
                                IYY = IYY
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function


'G02第二个象限的方案'
            Function G02XX2(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX + ddd, IYY), vbRed
                                IXX = IXX + ddd
                                IYY = IYY
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY + ddd), vbRed
                                IXX = IXX
                                IYY = IYY + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function
            

'G02第三个象限的方案'
            Function G02XX3(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX, IYY + ddd), vbRed
                                IXX = IXX
                                IYY = IYY + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX - ddd, IYY), vbRed
                                IXX = IXX - ddd
                                IYY = IYY
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function


'G02第四个象限的方案'
            Function G02XX4(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '定义起点X相对于圆心点的PXX
                Dim PYY As Double                          '定义起点Y相对于圆心点的PYY
                
                Dim Fmm As Double                          '定义判断Fmm
                Dim R2 As Double                           '定义圆弧半径的平方
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '计算圆弧半径的平方
                Dim bushu As Integer                           '设定记步数
                bushu = 1
                Do While bushu <= CCC  '判断插补是否结束
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX - ddd, IYY), vbRed
                                IXX = IXX - ddd
                                IYY = IYY
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY - ddd), vbRed
                                IXX = IXX
                                IYY = IYY - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '增加步数
                            End If
                        End If
                Loop
            End Function
































'该函数用于读取NC指令,并调用相应函数'
Function paint()
    Dim AnyLine() As String
    Dim EachLen As Integer '
    Dim Index As Integer
    For Index = 0 To length
        AnyLine = Split(NcLine(Index), Space(1))
        EachLen = UBound(AnyLine) - LBound(AnyLine)
        Dim Index2 As Integer
        For Each element In AnyLine
            Select Case Mid(element, 1, 1)
                Case "S"
                    S (element)
                Case "F"
                    F (element)
                Case "G"
                    G (element)
                    Case "X"
                    X (element)
                Case "Y"
                    Y (element)
                Case "M"
                    M (element)
                Case "I"
                    I (element)
                Case "J"
                    J (element)
            End Select
        Next element
        
        If judge = 0 Then                   '不必退出绘图函数，正常运行本绘图函数
            If InitF > 0 Then
                 Select Case InitG0123
                Case "G00"
                    G00 InitX, InitY, NextX, NextY
                Case "G01"
                    G01 InitX, InitY, NextX, NextY
                Case "G02"
                    G02 InitX, InitY, NextX, NextY, InitI, InitJ
                Case "G03"
                    G03 InitX, InitY, NextX, NextY, InitI, InitJ
            End Select
            End If
        Else                                          'M02或者M30指令
            Exit Function                               '退出本绘图函数
        End If
    Next Index
End Function

Private Sub Command1_Click()
    Picture1.Cls
End Sub

'用于单击‘读取文件’按钮时，读取相应路径下的NC代码文件并显示'
Private Sub DuQu_Click()
    StrFile = LuJin.Text
    intFile = FreeFile
    Open StrFile For Input As intFile
    strData = StrConv(InputB(FileLen(StrFile), intFile), vbUnicode)
    NCtext.Text = strData
    Close intFile
    NcLine = Split(strData, Chr(10))
    length = UBound(NcLine) - LBound(NcLine)
End Sub

'用于单击‘路径模拟’按钮时，初始化相关参量，并调用paint函数进行绘图'
Private Sub MoNi_Click()
    Picture1.Scale (-X0, Y0)-(16000 - X0, -10000 + Y0)
    Picture1.Line (-X0, 0)-(16000 - X0, 0)
    Picture1.Line (0, Y0)-(0, -10000 + Y0)
    Picture1.DrawWidth = 10
    Picture1.Circle (0, 0), 0, vbRed
    InitX = 0
    InitY = 0
    InitS = 0
    InitF = 0
    NextX = 0
    NextY = 0
    InitI = 0
    InitJ = 0
    judge = 0
    XAnytime.Text = InitX
    YAnytime.Text = InitY
    InitG0123 = "G00"
    InitG901 = "G90"
    Time = 1
    Picture1.DrawWidth = 2
    paint
    
End Sub

'定时器事件'
Private Sub Timer1_Timer()
    Time = 1
End Sub

