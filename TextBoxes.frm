VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Texbox Control For Beginners"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Click and then put mouse over the textbox"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Click to show an InputBox"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Click this to show a message box."
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Click this to make text in textbox uppercase"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Click this to change the background color"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click this button to make text colored"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "TextBoxes.frx":0000
      Top             =   480
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   360
      Top             =   1320
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Made by Mike
' purpose: to help beginers who are feeling lost and need
' to have help getting started
' if you have any questions , comments
' send email to mikesullins@diplomats.com
'thanks





Private Sub Command1_Click()
' the forecolor part means what color the text will
'be , there are many other colors you can use.
' Try puting these instead of vbred
'vbBlue
'vbWhite
'vbgreen
' and use RGB(10,80,120) try using other numbers in this too
Text1.ForeColor = vbRed ' try those on top and experiment
End Sub

Private Sub Command2_Click()
Text1.BackColor = vbRed
' the backcolor part means the color of
' the back of the text box. you can use any of the
' constants to change it
' constants in vb for colors are
' vbblack, vbred, vbblue, vbgreen, vbblack, and RGB(10,60,80)
' RGB stands for Red Green Blue, and in above example
' 10 is the amount of Red, 60 is the amount of Green
' and 80 is the amount of Blue

End Sub

Private Sub Command3_Click()
Text1.Text = UCase(Text1.Text)
' all this does is makes the text in the textbox
' become uppercase letters by using the UCase() function
' try using Text1.Text = Lcase(Text1.Text) and see what will
' happen.
End Sub

Private Sub Command4_Click()
Dim s As String ' here we make a varible named s that will hold a string ( a string is only a fancy way of saying text, or simpler terms , letters)
s = MsgBox("This is a message box.", vbOKOnly, "Hello")
' we say s is a message box using the MsgBox() function
' the "This is a message box" is what will be displayed in the message
' the vbOkOnly just says that it will have one button on the message, an Ok button
' the last part, "Hello" is what the title bar will say on the message box


End Sub

Private Sub Command5_Click()
Dim s As String ' make a varible called s that holds text
s = InputBox("This is an input box, please enter your name", "This is the title") ' the first part is whats in it, and the second is the title
MsgBox "Hello " & s ' this combines the msgbox() function in here to display the name entered in to the inputbox
End Sub

Private Sub Command6_Click()
Text1.ToolTipText = "Hello this is a tooltiptext"
' tooltiptext's are the bubbles that popup when the mouse rests
' on something. in other words just put the mouse
' pointer over the texbox after you click this button and wait till you see
' you cant miss it
End Sub
