VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  '���� ����
   Caption         =   "���� ������"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      BackColor       =   &H0000C000&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   300
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   9135
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   12135
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackColor       =   &H0000C000&
      BackStyle       =   0  '����
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   12135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private winhttp As New WinHttpRequest
Dim a
Dim aa As String
Dim b
Dim bb
Dim c
Dim d
Dim e
Dim f


Private Function Utf82String(ByRef data() As Byte) As String
Dim objStream
Dim strTmp As String
Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = "utf-8"
objStream.Mode = 3
objStream.Type = 1
objStream.Open
objStream.Write data
objStream.Flush
objStream.Position = 0
objStream.Type = 2
strTmp = objStream.ReadText
objStream.Close
Set objStream = Nothing
Utf82String = strTmp
End Function

Private Sub Command1_Click()
On Error GoTo error
f = 1
Label1.Caption = "������ �ε���"

With winhttp
    .Open "GET", "http://www.powersave.or.kr/f_img/queryPWR.aspx"
    .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:14.0) Gecko/20100101 Firefox/14.0"
    .Send
    .WaitForResponse
    Dim Temp$: Temp = Utf82String(.ResponseBody)
End With



   
   Text1.Text = Temp
a = Split(Split(Temp, "<data date=""")(1), """ time=")(0)
aa = Split(a, ".")(0) & "�� " & Split(a, ".")(1) & "�� " & Split(a, ".")(2) & "��"
'MsgBox aa
b = Split(Split(Temp, "time=""")(1), """ currentAmount=")(0)
bb = Split(b, ":")(0) & "�� " & Split(b, ":")(1) & "��"
'MsgBox bb
c = Split(Split(Temp, "currentAmount=""")(1), """ reserveAmount=")(0)

d = Split(Split(Temp, "reserveAmount=""")(1), """ reservePer=")(0)

e = Split(Split(Temp, "reservePer=""")(1), """ />")(0)

Label1.Caption = aa & " " & bb & " ���� ���� ���� : " & c & "�� kW ���� � ����� : " & d & "�� kW ���� � ������ " & e & "%"



If d < 100 Then
Label2.Caption = "�ɰ�"
Label2.BackColor = &HFF&
Me.BackColor = &HFF&
ElseIf d < 100 And d > 200 Then
Label2.Caption = "���"
Label2.BackColor = &H80FF&
Me.BackColor = &H80FF&
ElseIf d < 200 And d > 300 Then
Label2.Caption = "����"
Label2.BackColor = &HFFFF&
Me.BackColor = &HFFFF&

ElseIf d < 300 And d > 400 Then
Label2.Caption = "����"
Label2.BackColor = &H80FF80
Me.BackColor = &H80FF80
Else
Label2.Caption = "����"
Label2.BackColor = &HC000&
Me.BackColor = &HC000&
255
Exit Sub

error:
Label2.Caption = "����"
Label2.BackColor = &HFFFF&
Me.BackColor = &HFFFF&

End If




'Today.Caption = "���� " & Split(Split(Temp, "<div class=""anoday""><span class=""today"">Today <span class=""counterNum2"">")(1), "</span>")(0) & "�� �����߽��ϴ�"
'yesterday.Caption = "���� " & Split(Split(Temp, "<span class=""yesterday"">Yesterday <span class=""counterNum3"">")(1), "</span>")(0) & "�� �����߽��ϴ�"
'  Total.Caption = "��� " & Split(a, "</span>")(0) & "�� �����߽��ϴ�"

   ' MsgBox Split(Split(Temp, "</Marquee> </Marquee>")(0), "<Marquee behavior=alternate direction=up><Marquee behavior=alternate>")(1) ' ���ýô� ��û�մϴٶ���.

End Sub


Private Sub Form_Load()
Command1_Click
End Sub



Private Sub Label2_Click()


Command1_Click

End Sub

Private Sub Timer1_Timer()
If f >= 60 Then
f = 1
Command1_Click
Else
f = f + 1
End If
Text2.Text = f
End Sub
