VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���ز岹�������"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "�����Ļ"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "·��ģ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ȡ�ļ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ȴҺ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����״̬"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��������ϵ"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ʵʱ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����ٶ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����ת��"
      BeginProperty Font 
         Name            =   "����"
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
Dim InitX As Double '��ʼ��X'
Dim InitY As Double '��ʼ��Y'
Dim NextX As Double '��ֹ��X'
Dim NextY As Double '��ֹ��Y'
Dim InitS As Double '����ת��'
Dim InitF As Double '�����ٶ�'
Dim InitG901 As String '��Ի���������'
Dim InitG0123 As String 'G00��G01��G02��G03ָ��'
Dim InitI As Double 'X�������Բ��λ��'
Dim InitJ As Double 'Y�������Բ��λ��'
Dim Time As Integer
Dim judge As Integer



'Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)'

'�ú������ڻ�ȡ�����ٶ�InitS'
Function S(EachOn As String)
    S = Val(Mid(EachOn, 2))
    InitS = S
    ZhuZhou.Text = InitS
End Function

'�ú������ڻ�ȡ�����ٶ�InitF'
Function F(EachOn As String)
    F = Val(Mid(EachOn, 2))
    InitF = F
    Jinji.Text = InitF
End Function

'�ú������ڻ�ȡ�������һ����X�����������NextX'
Function X(EachOn As String)
    X = Val(Mid(EachOn, 2))
    If InitG901 = "G90" Then  'G90�»�ȡ��������'
        NextX = X
    Else
        NextX = X + InitX 'G91�»�ȡ�������'
    End If
End Function

'�ú������ڻ�ȡ�������һ����Y�����������NextY'
Function Y(EachOn As String)
    Y = Val(Mid(EachOn, 2))
    If InitG901 = "G90" Then 'G90�»�ȡ��������'
        NextY = Y
    Else
        NextY = Y + InitY 'G91�»�ȡ�������'
    End If
End Function

'�ú������ڻ�ȡԲ����X�����������'
Function I(EachOn As String)
    InitI = Val(Mid(EachOn, 2))
End Function

'�ú������ڻ�ȡԲ����Y�����������'
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

'�ú������ڻ�ȡ�������ȴҺ״̬'
'Function M(EachOn As String)
'    If Val(Mid(EachOn, 3)) = 3 Then
'        Motion.Text = "��ת"
'    ElseIf Val(Mid(EachOn, 3)) = 4 Then
'        Motion.Text = "��ת"
'    ElseIf Val(Mid(EachOn, 3)) = 5 Then
'        Motion.Text = "ͣת"
'    ElseIf Val(Mid(EachOn, 3)) = 7 Then
'        lqy.Text = "��"
'    ElseIf Val(Mid(EachOn, 3)) = 8 Then
'        lqy.Text = "��"
'    Else
'        lqy.Text = "�ر�"
'    End If
'End Function

Function M(EachOn As String)
    If Val(Mid(EachOn, 3)) = 3 Then         'M03ָ��
        Motion.Text = "��ת"
    ElseIf Val(Mid(EachOn, 3)) = 4 Then     'M04ָ��
        Motion.Text = "��ת"
    ElseIf Val(Mid(EachOn, 3)) = 5 Then     'M05ָ��
        Motion.Text = "ͣת"
    ElseIf Val(Mid(EachOn, 3)) = 7 Then     'M07ָ��
        lqy.Text = "��"
    ElseIf Val(Mid(EachOn, 3)) = 8 Then     'M08ָ��
        lqy.Text = "��"
    ElseIf Val(Mid(EachOn, 3)) = 9 Then     'M09ָ��
        lqy.Text = "��"
    ElseIf Val(Mid(EachOn, 3)) = 2 Then     'M02ָ��
        lqy.Text = "�ر�"
        Motion.Text = "ͣת"
        Call G00(InitX, InitY, 0, 0)
        judge = 1
    Else                                    'M30ָ��
        lqy.Text = "�ر�"
        Motion.Text = "ͣת"
        Call G00(InitX, InitY, X0, Y0)
        judge = 2
    End If
End Function

'�ú���Ϊ���ٽ���ָ��'
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

'�ú���ΰֱ�߲岹ָ��'
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


'�ú���Ϊ��Բ�岹ָ��'
Function G03(IX As Double, IY As Double, NX As Double, NY As Double, II As Double, JJ As Double)
    
    Dim deltaX As Double
    Dim deltaY As Double
    Dim dd As Double
    
    dd = InitF / 60                           '�������嵱��
                                        
    Dim Count As Integer
    
    Dim PX As Double                          '�������X�����Բ�ĵ��PX
    Dim PY As Double                          '�������Y�����Բ�ĵ��PY
    Dim PNX As Double                         '�����յ�X�����Բ�ĵ��PX
    Dim PNY As Double                         '�����յ�Y�����Բ�ĵ��PY
    
    Dim ZJDX As Double                         '�����м���ɵ�X��������
    Dim ZJDY As Double                         '�����м���ɵ�Y��������
    
    Dim RR2 As Double                           '����Բ���뾶��ƽ��
   
    
    II = II + IX                               '��Բ��I����������Ϊ��������
    JJ = JJ + IY                               '��Բ��J����������Ϊ��������
    RR2 = Sqr((IX - II) ^ 2 + (IY - JJ) ^ 2)         '����Բ���뾶
    
    PX = IX - II                              '��ʼ�������Բ���X����
    PY = IY - JJ                              '��ʼ�������Բ���Y����
    PNX = NX - II                             '��ֹ�������Բ���X����
    PNY = NY - JJ                             '��ֹ�������Բ���Y����
            
    Dim XiangXian As Integer                  '�����洢��ʼ�������
            
             '��ʼ���ڵ�һ���޵����'
            If PX > 0 And PY >= 0 And PNX >= 0 And PNY > 0 Then                    '��ʼ���ڵ�һ���޵��������ֹ���ڵ�һ����'
                XiangXian = 11
 
            ElseIf PX > 0 And PY >= 0 And PNX < 0 And PNY >= 0 Then                '��ʼ���ڵ�һ���޵��������ֹ���ڵڶ�����'
                XiangXian = 12
     
            ElseIf PX > 0 And PY >= 0 And PNX <= 0 And PNY < 0 Then                '��ʼ���ڵ�һ���޵��������ֹ���ڵ�������'
                XiangXian = 13
 
            ElseIf PX > 0 And PY >= 0 And PNX > 0 And PNY <= 0 Then                '��ʼ���ڵ�һ���޵��������ֹ���ڵ�������'
                XiangXian = 14
                
                
             '��ʼ���ڵڶ����޵����'
            ElseIf PX <= 0 And PY > 0 And PNX < 0 And PNY >= 0 Then               '��ʼ���ڵڶ����޵��������ֹ���ڵڶ�����'
                XiangXian = 21
            
            ElseIf PX <= 0 And PY > 0 And PNX <= 0 And PNY < 0 Then                '��ʼ���ڵڶ����޵��������ֹ���ڵ�������'
                XiangXian = 22
                
            ElseIf PX <= 0 And PY > 0 And PNX > 0 And PNY <= 0 Then                 '��ʼ���ڵڶ����޵��������ֹ���ڵ�������'
                XiangXian = 23
                
            ElseIf PX <= 0 And PY > 0 And PNX >= 0 And PNY > 0 Then                '��ʼ���ڵڶ����޵��������ֹ���ڵ�һ����'
                XiangXian = 24
            
            
            '��ʼ���ڵ������޵����'
            ElseIf PX < 0 And PY <= 0 And PNX <= 0 And PNY < 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 31
             
            ElseIf PX < 0 And PY <= 0 And PNX > 0 And PNY <= 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 32
                
            ElseIf PX < 0 And PY <= 0 And PNX >= 0 And PNY > 0 Then                 '��ʼ���ڵ������޵��������ֹ���ڵ�һ����'
                XiangXian = 33
                
            ElseIf PX < 0 And PY <= 0 And PNX < 0 And PNY >= 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵڶ�����'
                XiangXian = 34
               
               
            '��ʼ���ڵ������޵����'
            ElseIf PX >= 0 And PY < 0 And PNX > 0 And PNY <= 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 41
                
            ElseIf PX >= 0 And PY < 0 And PNX >= 0 And PNY > 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�һ����'
                XiangXian = 42
                
            ElseIf PX >= 0 And PY < 0 And PNX < 0 And PNY >= 0 Then                 '��ʼ���ڵ������޵��������ֹ���ڵڶ�����'
                XiangXian = 43
                
            ElseIf PX >= 0 And PY < 0 And PNX <= 0 And PNY < 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 44
 
            End If
            
            '������ʼ�����ֹ�������Բ�ĵ����ڵ�����ѡ��岹����'
                   Select Case XiangXian
                        '��һ����'
                        Case "11"
                            If IX <= NX Then
                                '��һ��Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan2 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                                '�����Բ��'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                Else
                                
                                '�ڶ����'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                End If
                                
                        Case "12"
                            '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                        Case "13"
                            '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                        Case "14"
                        '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                        
                        
                        '�ڶ�����'
                        Case "21"
                            If IX <= NX Then
                                '��һ��Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            
                            '�����Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            Else
                            
                            '�ڶ����'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "22"
                            '��һ��Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                        Case "23"
                            '��һ��Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                        Case "24"
                            '��һ��Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            
                        '��������'
                        Case "31"
                            If IX >= NX Then
                                '��һ��Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan2 IX, IY, II, JJ, dd, Count
                                '�����Բ��'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                            Else
                            
                                '�ڶ����'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "32"
                            '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                        Case "33"
                            '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                        Case "34"
                        '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            
                        
                        '���Ĵ���'
                        Case "41"
                            If IX >= NX Then
                                '��һ��Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan1 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan2 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan3 IX, IY, II, JJ, dd, Count
                                '�����Բ��'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G03fangan4 IX, IY, II, JJ, dd, Count
                            Else
                            '�ڶ����'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "42"
                            '��һ��Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                        Case "43"
                            '��һ��Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                        Case "44"
                        '��һ��Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan4 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan1 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan2 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G03fangan3 IX, IY, II, JJ, dd, Count
                    End Select
    
    InitX = NextX
    InitY = NextY
    InitI = 0                '��������ȱʡI0
    InitJ = 0                '��������ȱʡJ0
End Function
'G03��һ�����޵ķ���'
            Function G03fangan1(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX - ddd, IYY), vbRed
                                IXX = IXX - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY + ddd), vbRed
                                IYY = IYY + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function


'G03�ڶ������޵ķ���'
            Function G03fangan2(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX, IYY - ddd), vbRed
                                IYY = IYY - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX - ddd, IYY), vbRed
                                IXX = IXX - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function
            

'G03���������޵ķ���'
            Function G03fangan3(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX + ddd, IYY), vbRed
                                IXX = IXX + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY - ddd), vbRed
                                IYY = IYY - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function


'G03���ĸ����޵ķ���'
            Function G03fangan4(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
                    DoEvents
                        Fmm = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2 - R2
                        If Time Then
                            If Fmm >= 0 Then
                                Picture1.Line (IXX, IYY)-(IXX, IYY + ddd), vbRed
                                IYY = IYY + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX + ddd, IYY), vbRed
                                IXX = IXX + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function
 
 
 '�ú���Ϊ˳Բ�岹ָ��'
Function G02(IX As Double, IY As Double, NX As Double, NY As Double, II As Double, JJ As Double)
    
    Dim deltaX As Double
    Dim deltaY As Double
    Dim dd As Double
    
    dd = InitF / 60                           '�������嵱��
                                        
    Dim Count As Integer
    
    Dim PX As Double                          '�������X�����Բ�ĵ��PX
    Dim PY As Double                          '�������Y�����Բ�ĵ��PY
    Dim PNX As Double                         '�����յ�X�����Բ�ĵ��PX
    Dim PNY As Double                         '�����յ�Y�����Բ�ĵ��PY
    
    Dim ZJDX As Double                         '�����м���ɵ�X��������
    Dim ZJDY As Double                         '�����м���ɵ�Y��������
    
    Dim RR2 As Double                           '����Բ���뾶��ƽ��
   
    
    II = II + IX                               '��Բ��I����������Ϊ��������
    JJ = JJ + IY                               '��Բ��J����������Ϊ��������
    RR2 = Sqr((IX - II) ^ 2 + (IY - JJ) ^ 2)         '����Բ���뾶
    
    PX = IX - II                              '��ʼ�������Բ���X����
    PY = IY - JJ                              '��ʼ�������Բ���Y����
    PNX = NX - II                             '��ֹ�������Բ���X����
    PNY = NY - JJ                             '��ֹ�������Բ���Y����
            
    Dim XiangXian As Integer                  '�����洢��ʼ�������
            
             '��ʼ���ڵ�һ���޵����'
            If PX >= 0 And PY > 0 And PNX >= 0 And PNY > 0 Then                   '��ʼ���ڵ�һ���޵��������ֹ���ڵ�һ����'
                XiangXian = 11
 
            ElseIf PX >= 0 And PY > 0 And PNX < 0 And PNY >= 0 Then                '��ʼ���ڵ�һ���޵��������ֹ���ڵڶ�����'
                XiangXian = 12
     
            ElseIf PX >= 0 And PY > 0 And PNX <= 0 And PNY < 0 Then                '��ʼ���ڵ�һ���޵��������ֹ���ڵ�������'
                XiangXian = 13
 
            ElseIf PX >= 0 And PY > 0 And PNX > 0 And PNY <= 0 Then                '��ʼ���ڵ�һ���޵��������ֹ���ڵ�������'
                XiangXian = 14
                
                
             '��ʼ���ڵڶ����޵����'
            ElseIf PX <= 0 And PY >= 0 And PNX < 0 And PNY >= 0 Then              '��ʼ���ڵڶ����޵��������ֹ���ڵڶ�����'
                XiangXian = 21
            
            ElseIf PX < 0 And PY >= 0 And PNX <= 0 And PNY < 0 Then                '��ʼ���ڵڶ����޵��������ֹ���ڵ�������'
                XiangXian = 22
                
            ElseIf PX < 0 And PY >= 0 And PNX > 0 And PNY <= 0 Then                 '��ʼ���ڵڶ����޵��������ֹ���ڵ�������'
                XiangXian = 23
                
            ElseIf PX < 0 And PY >= 0 And PNX >= 0 And PNY > 0 Then                '��ʼ���ڵڶ����޵��������ֹ���ڵ�һ����'
                XiangXian = 24
            
            
            '��ʼ���ڵ������޵����'
            ElseIf PX <= 0 And PY < 0 And PNX <= 0 And PNY < 0 Then               '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 31
             
            ElseIf PX <= 0 And PY < 0 And PNX > 0 And PNY <= 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 32
                
            ElseIf PX <= 0 And PY < 0 And PNX >= 0 And PNY > 0 Then                 '��ʼ���ڵ������޵��������ֹ���ڵ�һ����'
                XiangXian = 33
                
            ElseIf PX <= 0 And PY < 0 And PNX < 0 And PNY >= 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵڶ�����'
                XiangXian = 34
               
               
            '��ʼ���ڵ������޵����'
            ElseIf PX > 0 And PY <= 0 And PNX > 0 And PNY <= 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 41
                
            ElseIf PX > 0 And PY <= 0 And PNX >= 0 And PNY > 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�һ����'
                XiangXian = 42
                
            ElseIf PX > 0 And PY <= 0 And PNX < 0 And PNY >= 0 Then                 '��ʼ���ڵ������޵��������ֹ���ڵڶ�����'
                XiangXian = 43
                
            ElseIf PX > 0 And PY <= 0 And PNX <= 0 And PNY < 0 Then                '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                XiangXian = 44
 
            End If
            
            '������ʼ�����ֹ�������Բ�ĵ����ڵ�����ѡ��岹����'
                   Select Case XiangXian
                        '��һ����'
                        Case "11"                                                 '��ʼ���ڵ�һ���޵��������ֹ���ڵ�һ����'
                            If IX >= NX Then
                                '��һ��Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '�����Բ��'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                            
                                Else
                                
                                '�ڶ����'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                            
                                End If
                                 
                        Case "12"                                                             '��ʼ���ڵ�һ���޵��������ֹ���ڵڶ�����'
                                '��һ��Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                             '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                        Case "13"                                                             '��ʼ���ڵ�һ���޵��������ֹ���ڵ�������'
                            '��һ��Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                        Case "14"                                                                '��ʼ���ڵ�һ���޵��������ֹ���ڵ�������'
                        '��һ��Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                        
                        
                        '�ڶ�����'
                        Case "21"                                                 '��ʼ���ڵڶ����޵��������ֹ���ڵڶ�����'
                            If IX >= NX Then
                                '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            ZJDX = II - RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                            
                            '�����Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            Else
                            
                            '�ڶ����'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "22"                                          '��ʼ���ڵڶ����޵��������ֹ���ڵ�������'
                            '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            ZJDX = II
                            ZJDY = JJ - RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                            '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                
                        Case "23"                                                  '��ʼ���ڵڶ����޵��������ֹ���ڵ�������'
                            '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            ZJDX = II + RR2
                            ZJDY = JJ
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                        Case "24"                                           '��ʼ���ڵڶ����޵��������ֹ���ڵ�һ����'
                            '��һ��Բ��'
                            ZJDX = II
                            ZJDY = JJ + RR2
                            deltaX = Abs(ZJDX - IX) / dd
                            deltaY = Abs(ZJDY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                            
                        '��������'
                        Case "31"                                            '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                            If IX <= NX Then
                                '��һ��Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '�����Բ��'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                            Else
                            
                                '�ڶ����'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "32"                                                     '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                            '��һ��Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                        Case "33"                                                    '��ʼ���ڵ������޵��������ֹ���ڵ�һ����'
                            '��һ��Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                        Case "34"                                                '��ʼ���ڵ������޵��������ֹ���ڵڶ�����'
                        '��һ��Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                            '�ڶ���Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                            
                        
                        '���Ĵ���'
                        Case "41"                                                         '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                            If IX <= NX Then
                                '��һ��Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                                ZJDX = II + RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX1 IX, IY, II, JJ, dd, Count
                                '�����Բ��'
                                deltaX = Abs(NX - IX) / dd
                                deltaY = Abs(NY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                            Else
                            '�ڶ����'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX4 IX, IY, II, JJ, dd, Count
                            End If
                            
                        Case "42"                                                            '��ʼ���ڵ������޵��������ֹ���ڵ�һ����'
                            '��һ��Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                                ZJDX = II
                                ZJDY = JJ + RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX2 IX, IY, II, JJ, dd, Count
                                '���Ķ�Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX1 IX, IY, II, JJ, dd, Count
                        Case "43"                                                          '��ʼ���ڵ������޵��������ֹ���ڵڶ�����'
                           '��һ��Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                                ZJDX = II - RR2
                                ZJDY = JJ
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX3 IX, IY, II, JJ, dd, Count
                                '������Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX2 IX, IY, II, JJ, dd, Count
                        Case "44"                                                    '��ʼ���ڵ������޵��������ֹ���ڵ�������'
                        '��һ��Բ��'
                                ZJDX = II
                                ZJDY = JJ - RR2
                                deltaX = Abs(ZJDX - IX) / dd
                                deltaY = Abs(ZJDY - IY) / dd
                                Count = deltaX + deltaY
                                G02XX4 IX, IY, II, JJ, dd, Count
                                '�ڶ���Բ��'
                            deltaX = Abs(NX - IX) / dd
                            deltaY = Abs(NY - IY) / dd
                            Count = deltaX + deltaY
                            G02XX3 IX, IY, II, JJ, dd, Count
                    End Select
    
    InitX = NextX
    InitY = NextY
    InitI = 0                'ȱʡI0
    InitJ = 0                'ȱʡJ0
End Function
'G02��һ�����޵ķ���'
            Function G02XX1(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
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
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX + ddd, IYY), vbRed
                                IXX = IXX + ddd
                                IYY = IYY
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function


'G02�ڶ������޵ķ���'
            Function G02XX2(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
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
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY + ddd), vbRed
                                IXX = IXX
                                IYY = IYY + ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function
            

'G02���������޵ķ���'
            Function G02XX3(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
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
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX - ddd, IYY), vbRed
                                IXX = IXX - ddd
                                IYY = IYY
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function


'G02���ĸ����޵ķ���'
            Function G02XX4(IXX As Double, IYY As Double, III As Double, JJJ As Double, ddd As Double, CCC As Integer)
                Timer1.Enabled = True
                Timer1.Interval = 10
                Dim PXX As Double                          '�������X�����Բ�ĵ��PXX
                Dim PYY As Double                          '�������Y�����Բ�ĵ��PYY
                
                Dim Fmm As Double                          '�����ж�Fmm
                Dim R2 As Double                           '����Բ���뾶��ƽ��
                R2 = (IXX - III) ^ 2 + (IYY - JJJ) ^ 2         '����Բ���뾶��ƽ��
                Dim bushu As Integer                           '�趨�ǲ���
                bushu = 1
                Do While bushu <= CCC  '�жϲ岹�Ƿ����
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
                                bushu = bushu + 1              '���Ӳ���
                            Else
                                Picture1.Line (IXX, IYY)-(IXX, IYY - ddd), vbRed
                                IXX = IXX
                                IYY = IYY - ddd
                                XAnytime.Text = IXX
                                YAnytime.Text = IYY
                                Time = 0
                                bushu = bushu + 1              '���Ӳ���
                            End If
                        End If
                Loop
            End Function
































'�ú������ڶ�ȡNCָ��,��������Ӧ����'
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
        
        If judge = 0 Then                   '�����˳���ͼ�������������б���ͼ����
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
        Else                                          'M02����M30ָ��
            Exit Function                               '�˳�����ͼ����
        End If
    Next Index
End Function

Private Sub Command1_Click()
    Picture1.Cls
End Sub

'���ڵ�������ȡ�ļ�����ťʱ����ȡ��Ӧ·���µ�NC�����ļ�����ʾ'
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

'���ڵ�����·��ģ�⡯��ťʱ����ʼ����ز�����������paint�������л�ͼ'
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

'��ʱ���¼�'
Private Sub Timer1_Timer()
    Time = 1
End Sub

