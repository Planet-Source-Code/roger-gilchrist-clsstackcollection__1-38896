VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ClsCollectionStack Demo"
   ClientHeight    =   8430
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Errors Demo"
      Height          =   375
      Index           =   19
      Left            =   1920
      TabIndex        =   35
      ToolTipText     =   "Force 4 Error messages from Function Exists"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RndMember (Remove) <="
      Height          =   615
      Index           =   18
      Left            =   2640
      TabIndex        =   32
      ToolTipText     =   "Move a random member of Stack 1 to 'Top' of Stack 2"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RndMember (Remove)   =>"
      Height          =   615
      Index           =   17
      Left            =   1320
      TabIndex        =   31
      ToolTipText     =   "Move a random member of Stack 1 to 'Top' of Stack 2"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "QuickSort =>"
      Height          =   375
      Index           =   16
      Left            =   2520
      TabIndex        =   30
      ToolTipText     =   "Sort Alphanumerically=>"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shuffle =>"
      Height          =   375
      Index           =   15
      Left            =   2520
      TabIndex        =   29
      ToolTipText     =   "Randomize collection=>"
      Top             =   5670
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<= QuickSort"
      Height          =   375
      Index           =   14
      Left            =   1320
      TabIndex        =   28
      ToolTipText     =   "<=Sort Alphanumerically"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<= Shuffle"
      Height          =   375
      Index           =   13
      Left            =   1320
      TabIndex        =   27
      ToolTipText     =   "<=Randomize collection"
      Top             =   5670
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Invert =>"
      Height          =   375
      Index           =   12
      Left            =   2520
      TabIndex        =   11
      ToolTipText     =   "Reverse order of collection=>"
      Top             =   5295
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert =>"
      Height          =   375
      Index           =   11
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "Text above inserted at selection point or 'Top'=>"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<= Insert"
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   8
      ToolTipText     =   "<=Text above inserted at selection point or 'Top'"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "v= Extract"
      Height          =   375
      Index           =   9
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Selected member moved to label below=>"
      Top             =   2415
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extract =v"
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "<=Selected member moved to label below"
      Top             =   2415
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Stacks"
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   1
      ToolTipText     =   "Empty both collections"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<= Invert"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "<=Reverse order of collection"
      Top             =   5295
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<= Pop "
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   3
      ToolTipText     =   "<=Top of stack moved to other stack"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<= Pull"
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   13
      ToolTipText     =   "<=Bottom of stack moved to other stack"
      Top             =   4905
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replace"
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   15
      ToolTipText     =   "Paste TextBox contents to stack and collection"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   6120
      Width           =   3135
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   3600
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pull =>"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Bottom of stack moved to other stack=>"
      Top             =   4905
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pop =>"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Top of stack moved to other stack=>"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Stacks"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Stack 1= full Stack 2= Empty"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   4
      Top             =   2895
      Width           =   1215
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottom"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   34
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Middle"
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   33
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Top"
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   26
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Min"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   25
      Top             =   1230
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Max"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   24
      Top             =   1005
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottom"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Middle"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   1740
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Top"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   1485
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Min"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1230
      Width           =   1815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Max"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   975
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   2805
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      Stack 1                                                         Stack 2"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"CollectionStack.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   6600
      Width           =   4695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu nuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuhelpopt 
         Caption         =   "clsStackCollection"
         Index           =   0
      End
      Begin VB.Menu mnuhelpopt 
         Caption         =   "ClsSafeCollection"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Plates1 As New ClsStackCollection
Public Plates2 As New ClsStackCollection
Private JustSayHello As New ClsSafeCollection
Private EditingList As Integer
Private Const StackHeight As Integer = 9
'StackHeight = number of plates pushed on if you want to change this don't forget to
'add to the Case 1 structure in Command1_Click and to make List1 and List2 taller to keep the illusion working
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Command1_Click(Index As Integer)

  Dim List1Target As Integer
  Dim List2Target As Integer

    List1Target = StackHeight - (List1.ListIndex) + 1
    List2Target = StackHeight - (List2.ListIndex) + 1
    '             MaxSizeofStack - ListIndex   + 1(for zero based lists)
    'This calculation is only required for the Demo
    'OR
    'if you are using a list or combo to view stack contents

    Select Case Index
      Case 0 'Reset Stacks
        Plates1.Clear
        Plates2.Clear
        '        With Plates1
        '            .Push PlateStr("A")
        '            .Push PlateStr("B")
        '            .Push PlateStr("C")
        '            .Push PlateStr("D")
        '            .Push PlateStr("E")
        '            .Push PlateStr("F")
        '            .Push PlateStr("G")
        '            .Push PlateStr("H")
        '            .Push PlateStr("I")
        '            .Push PlateStr("J")
        '        End With 'Plates1
        '
        ''This is equivalent to the With..End With structure commented out above
        Plates1.ArrayToCollection Array(PlateStr("A"), PlateStr("B"), PlateStr("C"), PlateStr("D"), PlateStr("E"), _
                                  PlateStr("F"), PlateStr("G"), PlateStr("H"), PlateStr("I"), PlateStr("J"))

      Case 1 'Pop =>
        Plates2.Push Plates1.Pop
      Case 2 'Pull =>
        Plates2.Push Plates1.Pull
      Case 3 'Replace
        If EditingList = 1 Then
            If List1.ListIndex > -1 Then
                Plates1.Replace List1Target, CVar(Text1.Text)
            End If
          Else 'NOT EDITINGLIST...
            If List2.ListIndex > -1 Then
                Plates2.Replace List2Target, CVar(Text1.Text)
            End If
        End If
        Text1.Text = ""
        Text1.Enabled = Len(Text1.Text)
      Case 4 '<= Pull
        Plates1.Push Plates2.Pull
      Case 5 '<= Pop
        Plates1.Push Plates2.Pop
      Case 6 '<=Invert
        Plates1.Invert
      Case 12 'Invert=>
        Plates2.Invert
      Case 7 'Clear Stacks
        Plates1.Clear
        Plates2.Clear
      Case 8 'Extract =>
        Label3.Caption = Plates1.Extract(List1Target)
      Case 9 '<= Extract
        Label3.Caption = Plates2.Extract(List2Target)
      Case 10 '<= Insert
        Plates1.Insert List1Target, Label3.Caption
        Label3.Caption = ""
      Case 11 'Insert =>
        Plates2.Insert List1Target, Label3.Caption
        Label3.Caption = ""
      Case 13 '<=Shuffle
        Plates1.Shuffle
      Case 14 '<=QuickSort
        Plates1.QuickSort
      Case 15 'Shuffle=>
        Plates2.Shuffle
      Case 16 'QuickSort=>
        Plates2.QuickSort
      Case 17 'RndMember (Remove)   =>
        Plates2.Push Plates1.RandomMember(True)
      Case 18 'RndMember (Remove)   <=
        Plates1.Push Plates2.RandomMember(True)
      Case 19
        Plates1.TriggerVerboseMsgBox
    End Select

    Stack List2, Plates2
    Stack List1, Plates1
    Label4(0).Caption = "Max:   " & Plates1.Max
    Label4(1).Caption = "Min:   " & Plates1.Min
    Label4(2).Caption = "Top:   " & Plates1.Top
    Label4(3).Caption = "Middle:" & Plates1.Middle
    Label4(4).Caption = "Bottom:" & Plates1.Bottom

    Label4(5).Caption = "Max:   " & Plates2.Max
    Label4(6).Caption = "Min:   " & Plates2.Min
    Label4(7).Caption = "Top:   " & Plates2.Top
    Label4(8).Caption = "Middle:" & Plates2.Middle
    Label4(9).Caption = "Bottom:" & Plates2.Bottom

    EnableButtons

End Sub

Private Sub EnableButtons()

    Command1(1).Enabled = Plates1.Count > 0 'Pop =>
    Command1(2).Enabled = Plates1.Count > 0 'Pull =>
    Command1(3).Enabled = Len(Text1.Text)   ' Replace
    Command1(4).Enabled = Plates2.Count > 0 '<= Pull
    Command1(5).Enabled = Plates2.Count > 0 '<= Pop

    Command1(7).Enabled = Plates1.Count > 0 'Clear Stacks

    Command1(8).Enabled = List1.ListIndex > -1 'Extract =>
    Command1(9).Enabled = List2.ListIndex > -1 '<= Extract
    Command1(10).Enabled = Len(Label3.Caption) '<= Insert
    Command1(11).Enabled = Len(Label3.Caption) 'Insert =>

    'These don't need to be active until there's more than one element
    'but the underlying routines can cope with zero and one member collections
    Command1(13).Enabled = Plates1.Count > 1 '<=Shuffle
    Command1(6).Enabled = Plates1.Count > 1  '<=Invert
    Command1(14).Enabled = Plates1.Count > 1 '<=QuickSort

    Command1(15).Enabled = Plates2.Count > 1 'Shuffle=>
    Command1(12).Enabled = Plates2.Count > 1 'Invert=>
    Command1(16).Enabled = Plates2.Count > 1 'QuickSort=>

    Command1(17).Enabled = Plates1.Count > 0 'RndMember (Remove)   =>
    Command1(18).Enabled = Plates2.Count > 0 'RndMember (Remove)   <=

End Sub

Private Sub Form_Initialize()
':) Ulli's VB Code Formatter V2.13.6 put this here but he is too modest
'to leave the message here when you rerun the Formatter.
'His program is a great addition to any coders utilities
'DOWNLOAD IT NOW AT www.planet-source-code.com!!!
'If you create a manifest file and are using XP then compiled program will
'use XP style for all controls not just form caption bars
    InitCommonControls

End Sub

Private Sub Form_Load()

    ListBase List1
    ListBase List2
    EnableButtons
    Command1_Click 0

End Sub

Private Sub List1_Click()

    EditingList = 1
    Command1(8).Enabled = List1.ListIndex > -1

    If List1.List(List1.ListIndex) <> "" Then
        Text1.Text = List1.List(List1.ListIndex)
        Command1(3).Enabled = Len(Text1.Text)
        Text1.Enabled = Len(Text1.Text)
    End If

End Sub

Private Sub List2_Click()

    EditingList = 2
    Command1(9).Enabled = List2.ListIndex > -1
    If List2.List(List2.ListIndex) <> "" Then
        Text1.Text = List2.List(List2.ListIndex)
        Command1(3).Enabled = Len(Text1.Text)
        Text1.Enabled = Len(Text1.Text)
    End If

End Sub

Private Sub ListBase(L As ListBox)

  Dim I As Integer

    L.Clear
    For I = 0 To StackHeight
        L.AddItem ""
    Next I

End Sub

Private Sub mnuFileExit_Click()

    End

End Sub

Private Sub mnuhelpopt_Click(Index As Integer)

    Select Case Index
      Case 0
        Plates1.About
      Case 1
        JustSayHello.About
    End Select

End Sub

Private Function PlateStr(ABC$) As String

  'DEMO ONLY
  'just to help the illusion along

    PlateStr = "\____" & ABC$ & "____/"

End Function

Private Sub Stack(L As ListBox, Coll As Variant)

  'DEMO ONLY
  'This is just fancy stuff to support the illusion of a stack of plates
  'it has nothing to do with the actual operation of the class
  
  Dim I As Integer
  Dim VisualPosition As Long ' used by the Edit window to get the correct member

    ListBase L
    VisualPosition = StackHeight
    For I = 1 To Coll.Count
        L.List(VisualPosition) = Coll.Item(I)
        VisualPosition = VisualPosition - 1
    Next I

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

  'DEMO ONLY

    If KeyCode = 13 Then 'why click a button when pressing Enter will do it anyway?
        KeyCode = 0
        Command1_Click 3
    End If

End Sub

':) Ulli's VB Code Formatter V2.13.6 (11/09/2002 10:26:27 AM) 9 + 242 = 251 Lines
