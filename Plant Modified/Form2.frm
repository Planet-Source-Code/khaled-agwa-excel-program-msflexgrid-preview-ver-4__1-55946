VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9570
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   706
      BackColor       =   8421504
      ForeColor       =   16711680
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Project"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CommonDialog1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "t2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Sheet 1"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid1(0)"
      Tab(1).Control(1)=   "Command1(0)"
      Tab(1).Control(2)=   "Command2(0)"
      Tab(1).Control(3)=   "Command3(0)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Sheet 2"
      TabPicture(2)   =   "Form2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid1(1)"
      Tab(2).Control(1)=   "Command3(1)"
      Tab(2).Control(2)=   "Command2(1)"
      Tab(2).Control(3)=   "Command1(1)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Sheet 3"
      TabPicture(3)   =   "Form2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Grid1(2)"
      Tab(3).Control(1)=   "Command3(2)"
      Tab(3).Control(2)=   "Command2(2)"
      Tab(3).Control(3)=   "Command1(2)"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "&Sheet 4"
      TabPicture(4)   =   "Form2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Grid1(3)"
      Tab(4).Control(1)=   "Command3(3)"
      Tab(4).Control(2)=   "Command2(3)"
      Tab(4).Control(3)=   "Command1(3)"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "&Sheet 5"
      TabPicture(5)   =   "Form2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid1(4)"
      Tab(5).Control(1)=   "Command3(4)"
      Tab(5).Control(2)=   "Command2(4)"
      Tab(5).Control(3)=   "Command1(4)"
      Tab(5).ControlCount=   4
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   375
         Index           =   4
         Left            =   -74640
         TabIndex        =   24
         ToolTipText     =   "Insert a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Index           =   4
         Left            =   -73800
         TabIndex        =   23
         ToolTipText     =   "Delete a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   375
         Index           =   4
         Left            =   -72840
         TabIndex        =   22
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   375
         Index           =   3
         Left            =   -74640
         TabIndex        =   20
         ToolTipText     =   "Insert a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Index           =   3
         Left            =   -73800
         TabIndex        =   19
         ToolTipText     =   "Delete a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   375
         Index           =   3
         Left            =   -72840
         TabIndex        =   18
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   375
         Index           =   2
         Left            =   -74640
         TabIndex        =   16
         ToolTipText     =   "Insert a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   -73800
         TabIndex        =   15
         ToolTipText     =   "Delete a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   375
         Index           =   2
         Left            =   -72840
         TabIndex        =   14
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   375
         Index           =   1
         Left            =   -74640
         TabIndex        =   12
         ToolTipText     =   "Insert a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Index           =   1
         Left            =   -73800
         TabIndex        =   11
         ToolTipText     =   "Delete a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   375
         Index           =   1
         Left            =   -72840
         TabIndex        =   10
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   375
         Index           =   0
         Left            =   -72960
         TabIndex        =   9
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Index           =   0
         Left            =   -73800
         TabIndex        =   4
         ToolTipText     =   "Delete a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   3
         ToolTipText     =   "Insert a row"
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox t2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Notice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1200
         TabIndex        =   1
         Top             =   1800
         Width           =   7455
         Begin VB.Label Label4 
            Caption         =   "Msflexgrid Preview Version 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            Top             =   1800
            Width           =   4095
         End
         Begin VB.Label Label3 
            Caption         =   "If my Code admire you,please Vote me,this will Support me much.Thanks alot for your Help and advice."
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   1080
            Width           =   6975
         End
         Begin VB.Label Label2 
            Caption         =   "I have spend alot of time in this project to put many Functions of Excel Sheet and i hope it admire you."
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   7095
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5535
         Index           =   0
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9763
         _Version        =   393216
         Rows            =   26
         Cols            =   10
         BackColor       =   16777215
         BackColorFixed  =   12632256
         ForeColorSel    =   8438015
         GridColor       =   0
         WordWrap        =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5535
         Index           =   1
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9763
         _Version        =   393216
         Rows            =   26
         Cols            =   10
         BackColor       =   16777215
         BackColorFixed  =   12632256
         ForeColorSel    =   8438015
         GridColor       =   0
         WordWrap        =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5535
         Index           =   2
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9763
         _Version        =   393216
         Rows            =   26
         Cols            =   10
         BackColor       =   16777215
         BackColorFixed  =   12632256
         ForeColorSel    =   8438015
         GridColor       =   0
         WordWrap        =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5535
         Index           =   3
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9763
         _Version        =   393216
         Rows            =   26
         Cols            =   10
         BackColor       =   16777215
         BackColorFixed  =   12632256
         ForeColorSel    =   8438015
         GridColor       =   0
         WordWrap        =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5535
         Index           =   4
         Left            =   -74760
         TabIndex        =   25
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9763
         _Version        =   393216
         Rows            =   26
         Cols            =   10
         BackColor       =   16777215
         BackColorFixed  =   12632256
         ForeColorSel    =   8438015
         GridColor       =   0
         WordWrap        =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "I Would like to thanks all people who helped me specially the users who are from vbcity.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1200
         TabIndex        =   27
         Top             =   5640
         Width           =   7455
      End
      Begin VB.Label Label1 
         Caption         =   "Vote Me"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         Top             =   4560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim KeyCode As Integer
Dim Index As Integer
Dim nAnswer As Integer
Dim temp As String, temp2 As Long
Dim j As Integer
Dim l As Integer
Dim i As Integer
Dim a33 As Boolean
Public selected As Integer
'*****************************************************

Public Sub Command1_Click(Index As Integer)
Grid1.Item(Index).AddItem (Grid1.Item(Index).Row), Grid1.Item(Index).Row
      
Header (Index)
End Sub
Public Sub Command2_Click(Index As Integer)
nAnswer = MsgBox("Do you want to delete this line", vbYesNo + vbQuestion, "Delete")
If nAnswer = vbYes Then
If Grid1.Item(Index).Rows > 2 Then
       Grid1.Item(Index).RemoveItem (Grid1.Item(Index).Row)
        Grid1.Item(Index).Rows = Grid1.Item(Index).Rows + 1
        Header (Index)
    End If
End If
  
End Sub
Public Sub Header(Index As Integer)
Grid1.Item(Index).TextMatrix(0, 0) = " "
End Sub

Public Sub Command3_Click(Index As Integer)
On Error GoTo errHandler
     With Form2.CommonDialog1
        .CancelError = True
        .Flags = cdlPDReturnIC Or cdlPDReturnDC
        .ShowPrinter
    End With
    
Printer.Orientation = 1
selected = Form2.SSTab1.Tab - 1
FlexGridPrint Form2.Grid1.Item(selected), , , , , 1
Exit Sub

errHandler:
'-----------
    MsgBox "Operation cancelled by user.", vbOKOnly, "Cancel"
End Sub

Private Sub Form_Load()
Me.Top = (MDIForm1.Height) * 0.0001
Me.Left = (MDIForm1.Width) * 0.0001
t2.Visible = False
For i = 0 To 4
Grid1.Item(i).ColWidth(0) = 0
Grid1.Item(i).RowHeight(0) = 0
Next i
End Sub

'******************************************************
'******************************************************
Private Sub Grid1_EnterCell(Index As Integer)
selected = Index + 1
'**********************************************************
  
t2.Visible = False
Grid1.Item(Index).CellBackColor = &HC0FFFF    'lt. yellow
Grid1.Item(Index).SetFocus
Grid1.Item(Index).Tag = Grid1.Item(Index)
End Sub

Private Sub Grid1_LeaveCell(Index As Integer)
    Grid1.Item(Index).CellBackColor = &H80000005  'white
End Sub

Private Sub Grid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 113        'F2
            Set_TextBox2 (Index)
    End Select
    
End Sub

Private Sub Grid1_KeyPress(Index As Integer, KeyAscii As Integer)

  '*****************************************************
        Select Case KeyAscii
        Case 13         'ENTER key
            KeyCode = 0
            INCR_CELL2 (Index)
        Case 8      'BkSpc
            Grid1.Item(Index) = Left$(Grid1.Item(Index), Len(Grid1.Item(Index)) - 1)
            Set_TextBox2 (Index)
        Case 27     'Esc - ignore
        Case Else
            Grid1.Item(Index) = Chr$(KeyAscii)
            t2 = Chr$(KeyAscii)
            Set_TextBox2 (Index)
    End Select
End Sub


Private Sub t2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27     'ESC - OOPS, restore old text
        t2 = Grid1.Item(Index).Tag
        t2.SelStart = Len(t2)
    Case 37     'Left Arrow
        If t2.SelStart = 0 And Grid1.Item(Index).Col > 1 Then
            Grid1.Item(Index).Col = Grid1.Item(Index).Col - 1
        Else
            If t2.SelStart = 0 And Grid1.Item(Index).Row > 1 Then
                Grid1.Item(Index).Row = Grid1.Item(Index).Row - 1
                Grid1.Item(Index).Col = Grid1.Item(Index).Cols - 1
            End If
        End If
    Case 38     'Up Arrow
        If Grid1.Item(Index).Row > 1 Then
            Grid1.Item(Index).Row = Grid1.Item(Index).Row - 1
        End If
    Case 39     'Rt Arrow
        If t2.SelStart = Len(t2) And Grid1.Item(Index).Col < Grid1.Item(Index).Cols - 1 Then
            Grid1.Item(Index).Col = Grid1.Item(Index).Col + 1
        Else
            If t2.SelStart = Len(t2) And Grid1.Item(Index).Row < Grid1.Item(Index).Rows - 1 Then
                Grid1.Item(Index).Row = Grid1.Item(Index).Row + 1
                Grid1.Item(Index).Col = 1
            End If
        End If
    Case 40     'Dn Arrow
        If Grid1.Item(Index).Row < Grid1.Item(Index).Rows - 1 Then
            Grid1.Item(Index).Row = Grid1.Item(Index).Row + 1
        End If
End Select
    IsCellVisible2 (Index)
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
    Dim pos%, l$, R$
    Select Case KeyAscii
        Case 13
              KeyAscii = 0
            Grid1.Item(selected - 1) = t2
            t2.Visible = False
            INCR_CELL2 (selected - 1)
            Grid1.Item(selected - 1).SetFocus
        Case 8                      'BkSpc - split string @ cursor
     
     
            pos% = t2.SelStart - 1 'where is the cursor?
            If pos% >= 0 Then
                l$ = Left$(Grid1.Item(selected - 1), pos%)       'left of cursor
                R$ = Right$(Grid1.Item(selected - 1), Len(Grid1.Item(selected - 1)) - pos% - 1) 'right of cursor
                Grid1.Item(selected - 1).Text = l$ + R$          'depleted string into Grid1.Item(Selected - 1)
            End If
        Case 27, 37 To 40
       
            Grid1.Item(selected - 1) = t2        'or it's going to look funny
        Case Else
       
            Grid1.Item(selected - 1) = t2 + Chr(KeyAscii)
    End Select
End Sub

Private Sub Grid1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Row%, Col%
t2.Visible = False
Row% = Grid1.Item(Index).MouseRow
Col% = Grid1.Item(Index).MouseCol
If Button = 2 And (Col% = 0 Or Row% = 0) Then
    Grid1.Item(Index).Col = IIf(Col% = 0, 1, Col%)    'rows?
    Grid1.Item(Index).Row = IIf(Row% = 0, 1, Row%)    'or cols?
    
End If
End Sub

Private Sub INCR_CELL2(Index As Integer)     'advance to next cell
  
    
    With Grid1.Item(Index)
    .HighLight = flexHighlightNever
    
    
If .Col < .Cols - 1 Then '*/*********change*******
.Col = .Col + 1
Else
If .Row < .Rows - 1 Then
.Row = .Row + 1                 'down 1 row
.Col = 1
End If
End If

    IsCellVisible2 (Index)
    .HighLight = flexHighlightAlways
    End With
End Sub
Private Sub Set_TextBox2(Index As Integer)   'put textbox over cell
    With t2
    .Top = Grid1.Item(Index).Top + Grid1.Item(Index).CellTop
    .Left = Grid1.Item(Index).Left + Grid1.Item(Index).CellLeft
    .Width = Grid1.Item(Index).CellWidth
    .Height = Grid1.Item(Index).CellHeight
    .Text = Grid1.Item(Index)
    .Visible = True
    .SelStart = Len(.Text)
    .SetFocus
    End With
End Sub
'this sub scrolls the cols / rows if they're not visible! (? why)
Private Sub IsCellVisible2(Index As Integer)
a33 = Grid1.Item(Index).CellTop
End Sub



