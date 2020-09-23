VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "KMH Software"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NEW"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PRINT"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "COPY"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CUT"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PASTE"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BLD"
            Object.ToolTipText     =   "Font"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LFT"
            Object.ToolTipText     =   "Left"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "JST"
            Object.ToolTipText     =   "Center"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RT"
            Object.ToolTipText     =   "Right"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlToolbarIcons 
      Left            =   1560
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":28B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":29C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3032
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3144
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3256
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New           "
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open...           "
         Shortcut        =   ^O
      End
      Begin VB.Menu sdfdss 
         Caption         =   "-"
      End
      Begin VB.Menu save 
         Caption         =   "&Save as           "
         Shortcut        =   ^S
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu SDS 
         Caption         =   "Page Setup.."
      End
      Begin VB.Menu print 
         Caption         =   "Print Current &Sheet"
      End
      Begin VB.Menu dfgd 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edite 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu clear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu selectall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu werwer 
         Caption         =   "-"
      End
      Begin VB.Menu addrow 
         Caption         =   "&Add Row"
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete Row"
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Begin VB.Menu write 
         Caption         =   "&Writing Font"
      End
      Begin VB.Menu Cell 
         Caption         =   "&Cell Font"
      End
   End
   Begin VB.Menu window 
      Caption         =   "&Window"
      Begin VB.Menu oo 
         Caption         =   "Data File"
      End
      Begin VB.Menu m1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu m2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu m3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu m4 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu m5 
         Caption         =   ""
         Visible         =   0   'False
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu topics 
         Caption         =   "&Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu yu 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'If you want to you my code,please let me know that!!!!
'Do not forget the copyright Laws
'Email: khaled.agwa@ gmx.net

'***********************************************
Option Explicit
Private srtRecentArray(5) As String
'***********************************************
Dim l As Integer
Dim i As Integer
Dim temp As String, temp2 As Long
'***************************************************
Dim OpenFile As String
Dim nAnswer As Integer
Dim Row, Rows As Integer
Dim Col, Cols As Integer
'*************************************************
Dim tmpText As String
Dim num As Integer
'**************************************************
Dim allCells As String
Dim fNum As Integer
Dim curRow, curCol As Integer
Dim FGrid() As MSFlexGrid

Private Sub about_Click()
MsgBox "KMH Software - by Khaled Agwa", vbApplicationModal + vbOKOnly, "About"
End Sub

Private Sub addrow_Click()
Form2.Command1_Click (Form2.SSTab1.Tab - 1)
End Sub

Private Sub Cell_Click()
On Error GoTo NocolorSelected
Form2.CommonDialog1.CancelError = True
Form2.CommonDialog1.Color = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellBackColor

Form2.CommonDialog1.ShowColor
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillRepeat
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellBackColor = Form2.CommonDialog1.Color
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillSingle

Exit Sub
NocolorSelected:
End Sub

Private Sub clear_Click()
Dim irow As Integer, icol As Integer, rSel As Integer, cSel As Integer, cRow As Integer, cCol As Integer
cCol = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Col
cRow = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Row
cSel = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).ColSel
rSel = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).RowSel
Dim tempq As Integer
If cRow > rSel Then
    tempq = rSel
    rSel = cRow
    cRow = tempq
End If
If cCol > cSel Then
    tempq = cSel
    cSel = cCol
    cCol = tempq
End If
For irow = cRow To rSel
    For icol = cCol To cSel
        Form2.Grid1.Item(Form2.SSTab1.Tab - 1).TextMatrix(irow, icol) = ""
    Next
Next

End Sub


Private Sub copy_Click()
tmpText = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Clip
Clipboard.clear
Clipboard.SetText tmpText
End Sub

Private Sub cut_Click()
tmpText = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Clip
Clipboard.clear
Clipboard.SetText tmpText
Dim irow As Integer, icol As Integer
For irow = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Row To Form2.Grid1.Item(Form2.SSTab1.Tab - 1).RowSel
For icol = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Col To Form2.Grid1.Item(Form2.SSTab1.Tab - 1).ColSel
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).TextMatrix(irow, icol) = ""
Next
Next
End Sub


Private Sub Exit_Click()
MsgBox "Please Vote me", vbOKOnly + vbInformation, "Thanks"
End
End Sub

Private Sub m1_Click()
'check if file exist
    If Dir(srtRecentArray(1)) = "" Then
        MsgBox ("file not found")
        srtRecentArray(1) = "" ' take it from the array
        Exit Sub
    End If
    ' your open actions here
Form2.CommonDialog1.FileName = srtRecentArray(1)
Call openx
End Sub
Private Sub m2_Click()
  'check if file exist
    If Dir(srtRecentArray(2)) = "" Then
        MsgBox ("file not found")
        srtRecentArray(2) = "" ' take it from the array
        Exit Sub
    End If
    ' your open actions here
Form2.CommonDialog1.FileName = srtRecentArray(2)
Call openx
End Sub
Private Sub m3_Click()
'check if file exist
    If Dir(srtRecentArray(3)) = "" Then
        MsgBox ("file not found")
        srtRecentArray(3) = "" ' take it from the array
        Exit Sub
    End If
    ' your open actions here
Form2.CommonDialog1.FileName = srtRecentArray(3)
Call openx
End Sub
Private Sub m4_Click()
 'check if file exist
    If Dir(srtRecentArray(4)) = "" Then
        MsgBox ("file not found")
        srtRecentArray(4) = "" ' take it from the array
        Exit Sub
    End If
    ' your open actions here
Form2.CommonDialog1.FileName = srtRecentArray(4)
Call openx
End Sub
Private Sub m5_Click()
'check if file exist
    If Dir(srtRecentArray(5)) = "" Then
        MsgBox ("file not found")
        srtRecentArray(5) = "" ' take it from the array
        Exit Sub
    End If
    ' your open actions here
Form2.CommonDialog1.FileName = srtRecentArray(5)
Call openx
End Sub

Private Sub MDIForm_Load()

' reset the menu
srtRecentArray(1) = ""
srtRecentArray(2) = ""
srtRecentArray(3) = ""
srtRecentArray(4) = ""
srtRecentArray(5) = ""

Call setLastFileMenu
'**************************************************

End Sub
Public Function setLastFileMenu()
'error trap
On Error GoTo err_setLastFileMenu
'Declarations
Dim i As Integer
Dim fn As Integer
'get the recent file information
'normaly you would store this in your DB or a seperate txtfile
'and put it in an array
'in this example it is in lastFiles.TXT
' intial settings
fn = FreeFile
i = 0

' read the last used files
Open App.path & "\lastFiles.txt" For Input As fn
While EOF(fn) = False And i < 5
    i = i + 1
    Input #fn, srtRecentArray(i)
Wend
'close the file
Close fn

' Load it in the menu caption
If srtRecentArray(1) <> "" Then
    Me.m1.Visible = False
    Me.m1.Caption = "&1. " & srtRecentArray(1)
    Me.m1.Visible = True
End If
If srtRecentArray(2) <> "" Then
    Me.m2.Caption = "&2. " & srtRecentArray(2)
    Me.m2.Visible = True
End If
If srtRecentArray(3) <> "" Then
    Me.m3.Caption = "&3. " & srtRecentArray(3)
    Me.m3.Visible = True
End If
If srtRecentArray(4) <> "" Then
    Me.m4.Caption = "&4. " & srtRecentArray(4)
    Me.m4.Visible = True
End If
If srtRecentArray(5) <> "" Then
    Me.m5.Caption = "&5. " & srtRecentArray(5)
    Me.m5.Visible = True
End If

Exit Function
err_setLastFileMenu:
MsgBox (Str(Err.Number) & "|" & Err.Description)
End Function
Public Function writeLastFileMenu(strLastFile As String)
'errortrap
On Error GoTo err_writeLastFileMenu
'Declarations
Dim i As Integer
Dim fn As Integer

srtRecentArray(0) = strLastFile

'Check if the entry already exist
For i = UBound(srtRecentArray) To 1 Step -1
     If srtRecentArray(i) = srtRecentArray(0) Then
     Exit Function
     End If
Next i

'reset the order
For i = UBound(srtRecentArray) To 1 Step -1
     srtRecentArray(i) = srtRecentArray(i - 1)
Next i

'rewrite the file
i = 0
fn = FreeFile

Open App.path & "\lastFiles.txt" For Output As fn
While i < 5
    i = i + 1
    Print #fn, srtRecentArray(i)
Wend
Close fn

'reset the menu
Call setLastFileMenu

Exit Function
err_writeLastFileMenu:
MsgBox (Str(Err.Number) & "|" & Err.Description)
End Function


Private Sub MDIForm_Unload(Cancel As Integer)
MsgBox "Please Vote me", vbOKOnly + vbInformation, "Thanks"
End Sub

Private Sub new_Click()
Form2.Show
End Sub

Private Sub open_Click()
On Error GoTo NoFileSelected

Form2.CommonDialog1.Filter = "KMH Software|*.KMH|All Files|*.*"
Form2.CommonDialog1.CancelError = True
Form2.CommonDialog1.InitDir = App.path
Form2.CommonDialog1.Flags = &H4
Form2.CommonDialog1.ShowOpen
If Form2.CommonDialog1.FileName = "" Then Exit Sub
Call openx
Exit Sub
NoFileSelected:
Exit Sub
End Sub


Private Sub paste_Click()
tmpText = Clipboard.GetText
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Clip = tmpText
End Sub

Private Sub print_Click()
Form2.Command3_Click (Form2.SSTab1.Tab - 1)
End Sub

Private Sub save_Click()
Call savex
End Sub

Private Sub SDS_Click()
On Error GoTo errHandler
     With Form2.CommonDialog1
        .CancelError = True
        .Flags = cdlPDReturnIC Or cdlPDReturnDC
        .ShowPrinter
    End With
 Exit Sub
      
errHandler:
'-----------
   ' MsgBox "Operation cancelled by user.", vbOKOnly, "Cancel"
End Sub

Private Sub selectall_Click()

Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Row = 1
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Col = 1
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).RowSel = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Rows - 1
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).ColSel = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).Cols - 1

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error Resume Next
Select Case Button.Key
'****************************
Case "COPY"
copy_Click
'*****************************
Case "CUT"
cut_Click
'*****************************
Case "NEW"
new_Click
'*****************************
Case "OPEN"
open_Click
'*****************************
Case "SAVE"
save_Click
'*****************************

Case "RT"
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillRepeat
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellAlignment = 7
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillSingle
'*****************************
Case "LFT"
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillRepeat
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellAlignment = 1
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillSingle

'*****************************
Case "BLD"
write_Click
'*****************************
Case "PASTE"
paste_Click
'*****************************
Case "PRINT"
print_Click
'*****************************
Case "JST"
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillRepeat
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellAlignment = 4
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillSingle
'*****************************

End Select
End Sub

Private Sub topics_Click()
MsgBox "No Help Topics avaliable now", vbOKOnly + vbExclamation, "Help"
End Sub

'*********************************************************
'*********************************************************
'*********************************************************

Public Sub openx()
On Error GoTo NoFileSelected

Dim FGrid(0 To 4) As MSFlexGrid 'Array To Be passed to SaveGrids
For i = 0 To 4
Set FGrid(i) = Form2.Grid1(i)
Next i

If Form2.CommonDialog1.FileName = "" Then Exit Sub

If LoadGrids(FGrid, Form2.CommonDialog1.FileName) Then
 '   MsgBox "Load Succesfull"
Else
 '   MsgBox "Load Unsuccesfull"
End If

Exit Sub

NoFileSelected:
Exit Sub
  
End Sub
Public Sub savex()
On Error GoTo NoFileSelected
Dim FGrid(0 To 4) As MSFlexGrid 'Array To Be passed to SaveGrids
For i = 0 To 4
Set FGrid(i) = Form2.Grid1(i)
Next i


    Form2.CommonDialog1.Filter = " KMH Software|*.KMH|" & "All Files|*.*"
    Form2.CommonDialog1.Flags = &H4
    Form2.CommonDialog1.DefaultExt = "KMH"
    Form2.CommonDialog1.FileName = ""
    Form2.CommonDialog1.Action = 2
       
If Form2.CommonDialog1.FileName = "" Then Exit Sub

If SaveGrids(FGrid, Form2.CommonDialog1.FileName) Then
   ' MsgBox "Save Succesfull"
Else
   ' MsgBox "Save Unsuccesfull"
End If
   Exit Sub
NoFileSelected:
    Exit Sub
  

   End Sub

'*****************************************************
'Name : SaveGrids, Type : Boolean, Returns : If the Save Was succesfull
'Author : Msg555
'
'Use : Save MSFlexGrids to be loaded with LoadGrids
'
'Parrameters
'   Flex() : Holds all of the Grids to be Saved
'   Path : Holds the file for information to Written to
'*****************************************************
Public Function SaveGrids(Flex() As MSFlexGrid, path As String) As Boolean
On Error GoTo e
Dim fNum As Long
fNum = FreeFile
Open path For Output As #fNum

'Loops through each FlexGrid in the array
Dim i As Integer 'index
Dim C As Integer 'Col
Dim R As Integer 'Row
For i = 0 To UBound(Flex)
    'You can add other info to be stored that vary from grid to grid
    Write #fNum, Flex(i).Cols, Flex(i).Rows, Flex(i).Row, Flex(i).Col
    For R = 0 To Flex(i).Rows - 1
        For C = 0 To Flex(i).Cols - 1
            'We have to select the Cell
            Flex(i).Col = C
            Flex(i).Row = R
            Write #fNum, Flex(i).TextMatrix(R, C), Flex(i).CellBackColor
        Next
    Next
Next

Close #fNum
SaveGrids = True
e:
End Function

'*****************************************************
'Name - LoadGrids, Type - Boolean, Return - If the Load Was succesfull
'Author - Msg555
'
'Use - Loads Files Saved using the SaveGrids function
'
'Parrameters
'   Flex() - Holds all of the Grids to be loaded
'   Path - Holds the file for information to inputed from
'*****************************************************
Public Function LoadGrids(Flex() As MSFlexGrid, path As String) As Boolean
On Error GoTo e
Dim fNum As Long
fNum = FreeFile
Open path For Input As #fNum

'Temporary Variables used to input values
Dim temp1 As String, temp2 As String, temp3 As String, temp4 As String, Ltemp As Long

Dim i As Integer 'index
Dim C As Integer 'Col
Dim R As Integer 'Row
'Loops through each FlexGrid in the array
For i = 0 To UBound(Flex)
    'You can add other info to be stored that vary from grid to grid
    Input #fNum, temp1, temp2, temp3, temp4
    Flex(i).Cols = temp1
    Flex(i).Rows = temp2
    
    For R = 0 To Flex(i).Rows - 1
        For C = 0 To Flex(i).Cols - 1
            'We have to select the Cell
            Flex(i).Col = C
            Flex(i).Row = R
            Input #fNum, temp1, Ltemp
            Flex(i).TextMatrix(R, C) = temp1
            Flex(i).CellBackColor = Ltemp
        Next
    Next
    'Selects the cell that was selected when grid was saved
    Flex(i).Row = temp3
    Flex(i).Col = temp4
Next

Close #fNum
LoadGrids = True
e:
End Function



Private Sub write_Click()
On Error GoTo NoFontSelected
Form2.CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
Form2.CommonDialog1.Color = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellForeColor
Form2.CommonDialog1.CancelError = True
Form2.CommonDialog1.FontName = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontName
Form2.CommonDialog1.FontBold = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontBold
Form2.CommonDialog1.FontItalic = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontItalic
Form2.CommonDialog1.FontSize = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontSize
Form2.CommonDialog1.Color = Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellForeColor
Form2.CommonDialog1.ShowFont
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillRepeat
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontName = Form2.CommonDialog1.FontName
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontBold = Form2.CommonDialog1.FontBold
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontItalic = Form2.CommonDialog1.FontItalic
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellFontSize = Form2.CommonDialog1.FontSize
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).CellForeColor = Form2.CommonDialog1.Color
Form2.Grid1.Item(Form2.SSTab1.Tab - 1).FillStyle = flexFillSingle

Exit Sub
NoFontSelected:
Exit Sub
End Sub
