VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   350
      SelectionMode   =   1
      FormatString    =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "andrew_bailon@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Command Button and Textbox with-IN the Flexgrid Control
'Feel free to modify it. You may freely distribute it
'without my prior notice: If ou want to chang to column position
'You may change it to any column you want
' I LOVE VISUAL BASIC: thanks to all my friends out there:


'PROGRAMMED BY: SIR ANDREW BAILON

Private Sub MoveTextBox()
     txtcell.Visible = True
      grid.Col = 0
      txtcell.Left = grid.Left + grid.CellLeft
      txtcell.Top = grid.Top + grid.CellTop
      txtcell.Height = grid.CellHeight
      txtcell.Width = grid.CellWidth
      txtcell.Text = grid.Text
End Sub
Private Sub MoveCommand()
      cmd.Visible = True
      grid.Col = 4  ' Change this if u want
      cmd.Left = grid.Left + grid.CellLeft
      cmd.Top = grid.Top + grid.CellTop
      cmd.Height = grid.CellHeight
      cmd.Width = grid.CellWidth
End Sub

Private Sub Form_Activate()
      MoveTextBox
End Sub

Private Sub Form_Load()
MoveCommand
grid.TextMatrix(1, 0) = "Sample"
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      MoveCommand
      MoveTextBox
End Sub

