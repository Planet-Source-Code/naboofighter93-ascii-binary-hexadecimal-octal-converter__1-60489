VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Binary|AscII|Hexadecimal|Octal  Converter"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "Binary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Boxes"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Decode"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decode"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7440
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdldialogbox 
      Left            =   2160
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox ot 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   3840
      MaxLength       =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "Binary.frx":0CCA
      Top             =   6240
      Width           =   7575
   End
   Begin VB.TextBox hex 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   3840
      MaxLength       =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Binary.frx":0CD2
      Top             =   4200
      Width           =   7455
   End
   Begin VB.TextBox bi 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   3840
      MaxLength       =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Binary.frx":0CE0
      Top             =   2280
      Width           =   7455
   End
   Begin VB.TextBox ac 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   3840
      MaxLength       =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Binary.frx":0CE9
      Top             =   360
      Width           =   7335
   End
   Begin VB.TextBox into 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6375
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   3720
      X2              =   12000
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   3720
      X2              =   11760
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   3720
      X2              =   11880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   3360
      X2              =   3360
      Y1              =   360
      Y2              =   7680
   End
   Begin VB.Label Label6 
      Caption         =   "Octal"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Hexadecimal"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label cl 
      Caption         =   "5000"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Characters left:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Binary"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "ASCII"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Text to convert:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Made by Miles
Dim t As Double
Dim j As String
Dim h(2) As String
Dim tot(2) As Double
Dim n As Double
Dim y(1 To 50000) As Double
Dim w(1 To 50000) As Double

Private Sub Command1_Click()
Dim t As String
Dim le As Integer
Dim a(5000) As String
Dim b(8) As Integer
Dim x(5000) As Double
Dim c As Integer
Dim d As Integer
Dim killer As String
Dim ans(3) As Double
Wait.Show
hex.Text = ""
ac.Text = ""
bi.Text = ""
ot.Text = ""
j = into.Text
n = 0
le = Len(j)
d = 0
Do
Wait.Label1.Caption = Wait.Label1.Caption & " "
'ascii convert
n = n + 1
a(n) = Mid(j, n, 1)
x(n) = Asc(a(n))
ac.Text = ac.Text & " " & x(n)
d = d + 1
temp = x(d)
k = ""
For c = 7 To 0 Step -1
'binary convert
If temp >= 2 ^ c Then
temp = temp - 2 ^ c
b(c) = 1
Else: b(c) = 0
End If
Next c
For c = 7 To 0 Step -1
'print binary
k = k & b(c)
z = b(c)
bi.Text = bi.Text & z
Next c
bi.Text = bi.Text & " "
h(1) = Left(k, 4)
h(2) = Right(k, 4)
'hex
For v = 1 To 2
tot(1) = 0
tot(2) = 0
If Mid(h(v), 4, 1) = "1" Then
tot(v) = tot(v) + 1
End If
If Mid(h(v), 3, 1) = "1" Then
tot(v) = tot(v) + 2
End If
If Mid(h(v), 2, 1) = "1" Then
tot(v) = tot(v) + 4
End If
If Mid(h(v), 1, 1) = "1" Then
tot(v) = tot(v) + 8
End If
m = tot(v)
If tot(v) <= 9 Then
killer = tot(v)
ElseIf m = 10 Then
killer = "A"
ElseIf m = 11 Then
killer = "B"
ElseIf m = 12 Then
killer = "C"
ElseIf m = 13 Then
killer = "D"
ElseIf m = 14 Then
killer = "E"
ElseIf m = 15 Then
killer = "F"
End If
If tot(v) > 9 Then
hex.Text = hex.Text & killer
Else
hex.Text = hex.Text & tot(v)
End If
Next v
hex.Text = hex.Text & " "
'Octal
oc = x(n) Mod 8
ans(1) = oc
ic = x(n) \ 8
zc = ic Mod 8
ans(2) = zc
ec = ic \ 8
uc = ec Mod 8
ans(3) = uc
ot.Text = ot.Text & " " & ans(3) & ans(2) & ans(1)
Loop While n < le
Wait.Label1.Caption = "Calculating, Please wait"
Wait.Hide
End Sub

Private Sub clear()
into.Text = ""
End Sub

Private Sub Command2_Click()
into.Text = ""
ac.Text = ""
j = bi.Text
le = Len(j)
n = 0
flag = 0
Do
n = n + 1
y(n) = Mid(j, n, 1)

If y(n) >= 2 Then
n = n - 1
MsgBox "Please enter only binary digits.", , "Error"
flag = 1
ElseIf Mid(j, n, 1) = " " Then
flag = 1
End If
Loop While n < le

If flag = 0 Then
For n = 1 To le
If Mid(j, n, 1) = 1 Then
temp = temp + Mid(j, n, 1) * 2 ^ (le - n)
End If
Next n

ac.Text = ac.Text & temp
gh = Chr(temp)
into.Text = gh

Call Command1_Click
End If
flag = 0
End Sub

Private Sub Command3_Click()
into.Text = ""
j = ac.Text
le = Len(j)

For n = 1 To le
If Mid(j, n, 1) <> " " Then
p = p & Mid(j, n, 1)
Else
into.Text = into.Text & Chr(p)
p = 0
End If
Next n
If p > 126 Then
MsgBox "Invalid numbers, enter a number 126 or below.", , "Invalid Characters"
Else
into.Text = into.Text & Chr(p)
Call Command1_Click
End If
End Sub

Private Sub Command4_Click()
into.Text = ""
bi.Text = ""
hex.Text = ""
ot.Text = ""
ac.Text = ""
End Sub

Private Sub Form_Load()
Wait.Hide
End Sub

Private Sub into_KeyPress(KeyAscii As Integer)
Dim cop As String
cop = Len(into.Text)
t = 1000
t = t - cop
cl.Caption = t
End Sub

Private Sub mnuexit_Click()
Unload Wait
Unload Me
End Sub

Private Sub mnuopen_Click()
Dim intp As String
Dim asci As String
Dim bin As String
Dim hexa As String
Dim octa As String

Dim intfilenum As Integer, strtextline As String
cdldialogbox.InitDir = "C:\"
cdldialogbox.Filter = "*.txt"
cdldialogbox.Flags = cdlOFNFileMustExist
cdldialogbox.ShowOpen

intfilenum = FreeFile
Open cdldialogbox.FileName For Input As #intfilenum
Line Input #intfilenum, intp
Line Input #intfilenum, asci
Line Input #intfilenum, bin
Line Input #intfilenum, hexa
Line Input #intfilenum, octa
Close #intfilenum
into.Text = intp
ac.Text = asci
bi.Text = bin
hex.Text = hexa
ot.Text = octa
End Sub

Private Sub mnusave_Click()
Dim intfilenum As Integer
intfilenum = FreeFile
cdldialogbox.InitDir = "C:\"
cdldialogbox.Filter = ""
cdldialogbox.Flags = cdlOFNOverwritePrompt
cdldialogbox.ShowSave
Open cdldialogbox.FileName For Output As #intfilenum
Print #intfilenum, into.Text
Print #intfilenum, ac.Text
Print #intfilenum, bi.Text
Print #intfilenum, hex.Text
Print #intfilenum, ot.Text
Close #intfilenum
End Sub
