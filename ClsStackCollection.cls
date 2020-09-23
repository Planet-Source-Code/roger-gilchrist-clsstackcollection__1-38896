VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsStackCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Copyright 2002 Roger Gilchrist
'I started developing this class on the fly when I needed a LIFO (Last In First Out) stack
'then expanded it to include other Methods and Properties
'I don't know if this is a re-invention of the wheel or not but I hope you like it.

'It is inspired by ClsStack in 'Visual Basic 4 SECRETS' by Harold Davis from IDG Books

'RECOMMENDATION 1: When you're developing something read over old, even out-of-date, books and articles,
'                  you never know what might suddenly become meaningful or at least suggest a new approach.
'RECOMMENDATION 2: Do this while you are away from your computer, it gives you time to think without being
'                  able to get lost in the act of coding.

'Search for '*---PROGRAMMER MODIFICATION POINT---*
'to see where you might like to modify the class for release purposes

'STACK STUFF
' Create a collection of 5 items
'           Push A
'           Push B
'           Push C
'           Push D
'           Push E
' imagine them as a pile of plates
' top      \____E____/
'          \____D____/
'          \____C____/
'          \____D____/
' bottom   \____A____/
'a LIFO stack would Pop E off the top
'a FIFO stack would Pull A from the bottom

'This class adds several new Methods and Properties to collections
'To use class just put a declaration on  your form or module
' i.e. Public Plates1 As New ClsCollection
'NEW METHODS
'---------------Push and Pop are what I needed for creating a LIFO (Last In First Out) stack
'---------------The rest is just stuff I thought would make the class more useful
'ArrayFromCollection:   Put contents of collection into an array(used internally by quckSort and Shuffle)
'ArrayToCollection:     Load members of an array into the class/collection(used internally by quckSort and Shuffle)
'Bottom:        Returns first value added to collection
'Clear:         Empty collection in one call
'Exists:        Tests if index/Key is in collection. Uses error trap to test. Optional 'Silent' allows internal calls to silence verbose
'Extract:       Remove any member from collection
'Invert:        Reverse order of collection
'Max:           Returns the largest (or lowest if alphanumeric) member of collection
'Min:           Returns the smallest (or highest if alphanumeric) member of collection
'Pop:           Use collection as a LIFO (Last In First Out) stack, removes member from collection.
'Pull:          Use collection as a FIFO (First In First Out) stack, removes member from collection. (named as opposite to Push)
'Push:          Wrapper for VB standard Add (Push does not need the testing overheads of my modified Add method).
'               Push is used by both LIFO & FIFO stacks.
'               (Your code is more readable if you use Push|Pop or Push|Pull rather than Add and Remove
'QuickSort:     Sort collection into alphanumeric order
'RandomMember:  Extract a random member of the collection (Optional RemoveIt as Boolean = False); if True remove from collection
'               (if your collection is a pack of cards you need to remove it
'               but if it is a coin-toss you would leave it in the collection)
'Replace:       Change the Item for an Index/Key
'Shuffle:       Randomise the contents of the array
'Top:           Returns last value added to collection
'Notes: QuickSort, Shuffle and Replace are slow for very large collections
'       as they relay on reading the collection into an array (another collection in Replace)

'INTERNAL VARIABLES 'Set Verbose and AutoSafety from Class_Initialize
'Verbose: If True then Show MsgBox Error comment and offer to override error; False No MsgBox.
'AutoSafety: If True then if there is an automatic way to override the error take it (equals Yes in Verbose mode)
'Notes: If Verbose is True then AutoSafety will never trigger no matter what its value
'       'as is' the code uses IsDebugMode to turn Verbose On and AutoSafety Off if you are in the IDE
'       You should set Verbose to False in released programs
'       AutoSafety is a matter of taste; what do you want to happen if bad data reaches this class?
'       Verbose is useful for debugging your code as it will warn you that something unexpected has happened
'       if you like what the Verbose version does then set AutoSafety to True in Released program.

Private m_Collection As New Collection
Private Verbose As Boolean
Private AutoSafety As Boolean
Private Const Silent As Boolean = True

Public Sub About()

  'Copyright 2002 Roger Gilchrist
  'NOTE VB help comments "The maximum length of prompt is approximately 1024 characters,
  '                        depending on the width of the characters used."
  'This is why I have used vbCr rather than vbNewLine it saves one character per line.
  'This message is very close to the limit mentioned.
  'It is also at the maximum number of line continuations

  Dim msg As String

    msg = "Use VB's collection as Stacks:" & vbCr & "Call" & vbTab & vbTab & vbTab & "Action" & vbCr & _
          "-----------------------------------------------------------" & vbCr & _
          "ArrayFromCollection" & vbTab & "Put stack into an array" & vbCr & _
          "ArrayToCollection:" & vbTab & vbTab & "Load an array into the stack" & vbCr & _
          "Clear" & vbTab & vbTab & vbTab & "Empty the stack" & vbCr & _
          "Exists IndexKey" & vbTab & vbTab & "Test member exists" & vbCr & _
          "Extract(IndexKey)" & vbTab & vbTab & "Get and remove a member" & vbCr & _
          "Invert" & vbTab & vbTab & vbTab & "Invert stack" & vbCr & _
          "Insert IndexKey, NewValue" & vbTab & "Place value at point in stack" & vbCr & _
          "Pop" & vbTab & vbTab & vbTab & "Act as Last-In-First-Out (LIFO) stack" & vbCr & _
          "Push NewValue" & vbTab & vbTab & "Add to LIFO & FIFO stacks" & vbCr & _
          "Pull" & vbTab & vbTab & vbTab & "Act as First-In-First-Out (FIFO) stack" & vbCr & _
          "QuickSort" & vbTab & vbTab & vbTab & "Alphanumeric sort" & vbCr & _
          "RandomMember(RemoveIt)" & vbTab & "Get random member (Optional remove it)" & vbCr & _
          "Replace IndexKey, NewValue" & vbTab & "Edit a stack member" & vbCr & _
          "Shuffle" & vbTab & vbTab & vbTab & "Randomise the stack" & vbCr & _
          "Min" & vbTab & vbTab & vbTab & "Lowest alphanumeric member" & vbCr & _
          "Max" & vbTab & vbTab & vbTab & "Highest alphanumeric member" & vbCr & _
          "Top" & vbTab & vbTab & vbTab & "Last member added" & vbCr & _
          "Middle" & vbTab & vbTab & vbTab & "Middle member(or 1 less if Count is even)" & vbCr & _
          "Bottom" & vbTab & vbTab & vbTab & "First member added" & vbCr & _
          "-----------------------------------------------------------" & vbCr & _
          vbCr & _
          "Hope you like it!" & vbCr & _
          "Copyright 2002 Roger Gilchrist"
    MsgBox msg, vbInformation, "ClsStackCollection"

End Sub

Public Function ArrayFromCollection() As Variant

  'Copyright 2002 Roger Gilchrist
  'Extract Collection into an array
  'Used internally by QuickSort and Shuffle

  Dim Tmpa As Variant, I As Long

    ReDim Tmpa(m_Collection.Count - 1) As Variant
    For I = 1 To m_Collection.Count
        Tmpa(I - 1) = m_Collection.Item(I)
    Next I
    ArrayFromCollection = Tmpa

End Function

Public Sub ArrayToCollection(LArray As Variant)

  'Copyright 2002 Roger Gilchrist
  'Load an Array into the class/collection
  'Used internally by QuickSort and Shuffle
  'Because you can't directly edit collections
  'the routine creates a temporary collection
  'then overwrites the class collection using Set

  Dim TmpColl As New Collection, I As Long

    For I = LBound(LArray) To UBound(LArray)
        TmpColl.Add LArray(I)
    Next I
    Set m_Collection = TmpColl

End Sub

Public Property Get Bottom() As Variant

  'Copyright 2002 Roger Gilchrist
  'Return largest member of collection
  'if numeric then numeric
  'if alpha or alphanumeric then lowest in alpha sort

    If Exists(1, Silent) Then
        Bottom = m_Collection(1)
    End If

End Property

Private Sub Class_Initialize()

  '*---PROGRAMMER MODIFICATION POINT---*
  'Change Verbose to False 1. if you want to test the AutoSafety error handling
  '                        2. You probably don't want end-users to see the messages
  '
  'AutoSafety: Depending on how you want to handle errors you might want to set this to False
  '            if you don't like the classes default corrections or change the defaults

    Verbose = IsDebugMode     ' Use MsgBox to warn of error and offer a way out
    AutoSafety = Not Verbose ' Take the way out if it exists

    ' NOTE: If Verbose is True then AutoSafety will never trigger no matter what its value

End Sub

Private Sub Class_Terminate()

  'Free up memory

    Clear

End Sub

Public Sub Clear()

  'Copyright 2002 Roger Gilchrist
  'wipe collection from memory

    Set m_Collection = Nothing

End Sub

Public Property Get Count() As Variant

  'Wrapper for VB standard Collection.Count

    Count = m_Collection.Count

End Property

Public Sub TriggerVerboseMsgBox()

    If Exists(-1) Then
    End If
    If Exists(Count + 1) Then
    End If
    If Exists("Wally") Then
    End If
    If Not Exists("Wally", Silent) Then
        MsgBox "Exists failed but was set to Silent", vbInformation, "Errors Demo"
    End If

End Sub

Public Property Get Exists(IndexKey As Variant, Optional Quiet As Boolean = False) As Boolean

  'Guard routine for other routines allows them to fail gracefully
  'Quiet allow some Properties to fail without triggering the Verbose MsgBox even when it is turned on

  Dim msg As String ' junk value: routine doesn't care what it is, just whether or not it can be set

    'Verbose reuses it
    On Error GoTo NotExists
    msg = m_Collection.Item(IndexKey)
    Exists = True

Exit Property

NotExists:
    If Not Quiet Then
        If Verbose Then
            msg = ""
            If IsNumeric(IndexKey) Then
                If IndexKey < 1 Then
                    msg = msg & "Index: '" & IndexKey & "' is below the minimum Index value of 1."
                  ElseIf IndexKey > m_Collection.Count Then 'NOT INDEXKEY...
                    msg = msg & "Index: '" & IndexKey & "' is above the collection's Count of " & m_Collection.Count
                End If
              Else 'ISNUMERIC(INDEXKEY) = FALSE
                msg = msg & "Key: '" & IndexKey & "' is not in the collection."
            End If
            MsgBox msg, vbCritical, "ClsStackCollection Error"
        End If
    End If

End Property

Public Function Extract(IndexKey As Variant) As Variant

  'Copyright 2002 Roger Gilchrist
  'Get IndexKey member and remove it from the collection

    If Exists(IndexKey, Silent) Then
        Extract = m_Collection.Item(IndexKey)
        m_Collection.Remove IndexKey
    End If

End Function

Public Sub Insert(IndexKey As Variant, InsertValue As Variant)

  'Copyright 2002 Roger Gilchrist
  'Put InsertValue into the collection at a specific point

  Dim m_Edit As New Collection 'Tmp collection
  Dim Inserted As Boolean
  Dim I As Long

    For I = 1 To m_Collection.Count
        If I <> IndexKey Then
            m_Edit.Add m_Collection.Item(I)
          Else 'NOT I...
            m_Edit.Add InsertValue
            m_Edit.Add m_Collection.Item(I)

            Inserted = True
        End If
    Next I
    If Inserted = False Then ' copes with 0 members before insertion
        m_Edit.Add InsertValue
    End If
    Set m_Collection = m_Edit

End Sub

Public Sub Invert()

  'Copyright 2002 Roger Gilchrist
  'Reverse the order of an array

  Dim m_Edit As New Collection 'Tmp collection
  Dim I As Long

    For I = m_Collection.Count To 1 Step -1
        m_Edit.Add m_Collection.Item(I)
    Next I
    Set m_Collection = m_Edit

End Sub

Private Function IsDebugMode(Optional bSetMode As Boolean = False) As Boolean

  'VB2MAX 'Tip of the Week: Check Whether VB Is in Debug Mode
  'Erik Perrohe (Seattle, WA)
  '*---PROGRAMMER MODIFICATION POINT---*
  'For Final Release versions remove this and the call to it
  'in Class_Initilaize

  Static DebugMode As Boolean

    DebugMode = bSetMode
    If Not DebugMode Then
        Debug.Assert IsDebugMode(True)
    End If
    IsDebugMode = DebugMode

End Function

Public Property Get Item(IndexKey As Variant) As Variant

  'Copyright 2002 Roger Gilchrist
  'Wrapper for standard Item but with added safety returns an empty Variant on error

    If Exists(IndexKey) Then
        Item = m_Collection.Item(IndexKey)
      Else 'EXISTS(INDEXKEY) = FALSE
        If Verbose Then
            MsgBox "Your Index/Key: '" & IndexKey & "' does not exist.", vbOKOnly + vbInformation, "Item does not exist."
          ElseIf AutoSafety Then 'VERBOSE = FALSE
            '*---PROGRAMMER MODIFICATION POINT---*
            'there is no AutoSafety for this condition
            'you might like to supply one
            ' an empty Variant is returned
        End If
    End If

End Property

Public Property Get Max() As Variant

  'Copyright 2002 Roger Gilchrist
  'Return largest member of collection
  'if numeric then numeric
  'if alpha or alphanumeric then lowest in alpha sort

  Dim Member As Variant

    For Each Member In m_Collection
        If Member > Max Then
            Max = Member
        End If
    Next Member

End Property

Public Property Get Middle() As Variant

  'Copyright 2002 Roger Gilchrist
  'I can't think of any use for it but who knows
  'Return middle member of collection
  '*---PROGRAMMER MODIFICATION POINT---*
  'You could modify this to get other fractional members i.e. quarter points
  'if Count is odd its the real middle
  'Else the item 1 below middle (see below for other possible returns

  Dim Member As Variant, MidPoint As Long

    If m_Collection.Count = 1 Then
        MidPoint = m_Collection.Count
      Else 'NOT M_COLLECTION.COUNT...
        If m_Collection.Count Mod 2 <> 0 Then '
            MidPoint = (m_Collection.Count + 1) \ 2
          Else 'NOT M_COLLECTION.COUNT...
            MidPoint = (m_Collection.Count) \ 2
            '*---PROGRAMMER MODIFICATION POINT---*
            'Exit Property                           'Uncomment if you want a blank
            'MidPoint = (m_Collection.Count) \ 2 + 1 'Uncomment for I above mid point

        End If
    End If
    If Exists(MidPoint, Silent) Then
        Middle = Item(MidPoint)
    End If

End Property

Public Property Get Min() As Variant

  'Copyright 2002 Roger Gilchrist
  'Return Smallest member of collection
  'if numeric then numeric
  'if alphanumeric then highest in alpha sort

  Dim Member As Variant

    Min = Max 'force min to highest value so that it will hit in the If ...Then structure
    For Each Member In m_Collection
        If Member < Min Then
            Min = Member
        End If
    Next Member

End Property

Public Function Pop() As Variant

  'Copyright 2002 Roger Gilchrist
  'Return last member of collection and remove it from collection
  ' used to get and remove 'top' of a LIFO (Last In First Out) stack
  'top and bottom are the same last member but seen from the pespective of your stack(see diagram at top of code)

    With m_Collection
        If .Count Then
            Pop = m_Collection(.Count)
            .Remove .Count
        End If
    End With 'M_COLLECTION

End Function

Public Function Pull() As Variant

  'Copyright 2002 Roger Gilchrist
  'Return First member of collection and remove it from collection
  'used to get and remove 'bottom' of a FIFO (First In First Out) stack
  'top and bottom are the same last member but seen from the pespective of your stack(see diagram at top of code)

    If Exists(1, Silent) Then ' Remembering collections are 1 based and self organizing
        Pull = m_Collection(1)
        m_Collection.Remove 1
    End If

End Function

Public Sub Push(Item As Variant)

  'Copyright 2002 Roger Gilchrist
  'Wrapper for VB standard Collection.Add
  ' used to add to 'top' of a LIFO (Last In First Out) stack
  ' used to add to 'bottom' of a FIFO (First In First Out) stack
  'top and bottom are the same last member but seen from the pespective of your stack(see diagram at top of code)

    m_Collection.Add Item

End Sub

Private Sub QSort(LArray As Variant, L As Variant, R As Variant)

  'Performs a QuickSort

  Dim I As Variant, J As Variant, X As Variant, Y As Variant

    I = L
    J = R
    X = LArray((L + R) / 2)
    While (I <= J)
        While (LArray(I) < X And I < R)
            I = I + 1
        Wend
        While (X < LArray(J) And J > L)
            J = J - 1
        Wend
        If (I <= J) Then
            SwapAnyThing LArray(I), LArray(J)
            I = I + 1
            J = J - 1
        End If
    Wend
    If (L < J) Then
        QSort LArray, L, J
    End If
    If (I < R) Then
        QSort LArray, I, R
    End If

End Sub

Public Sub QuickSort()

  'Copyright 2002 Roger Gilchrist
  'An Alphanumeric Sort of the collection
  
  Dim Tmpa As Variant

    Tmpa = ArrayFromCollection
    QSort Tmpa, LBound(Tmpa), UBound(Tmpa)
    ArrayToCollection Tmpa

End Sub

Public Function RandomMember(Optional RemoveIt As Boolean = False) As Variant

  'Copyright 2002 Roger Gilchrist
  'Get a random member of the collection and optionally remove it from the collection
  'If your collection is:
  'a pack of cards then you would remove it so it could not be redrawn
  'a coin-toss then you leave it in

  Dim RndMember As Long

    RndMember = Int(Rnd * Count + 1)
    RandomMember = m_Collection.Item(RndMember)
    If RemoveIt Then
        m_Collection.Remove RndMember
    End If

End Function

Public Sub Replace(IndexKey As Variant, NewValue As Variant)

  'Copyright 2002 Roger Gilchrist
  'Works by creating a temporary copy of m_collection but replacing the IndexKey value with your NewValue
  'because it has to work through the whole collection it will be slow for large collections

  Dim I As Long
  Dim m_Edit As New Collection 'Tmp collection

    If Exists(IndexKey) Then
        For I = 1 To m_Collection.Count
            If I <> IndexKey Then
                m_Edit.Add m_Collection.Item(I)
              Else 'NOT I...
                m_Edit.Add NewValue
            End If
        Next I
        Set m_Collection = m_Edit
      Else 'EXISTS(INDEXKEY) = FALSE
        If AutoSafety Then 'VERBOSE = FALSE
            '*---PROGRAMMER MODIFICATION POINT---*
            'there is no AutoSafety for this condition
            'you might like to supply one
        End If
    End If

End Sub

Public Sub Shuffle()

  'Copyright 2002 Roger Gilchrist
  'Randomise order in the collection
  'The  For...Next structure is simply to give a number of hits on SwapAnyThing

  Dim Tmpa As Variant, I As Long

    Tmpa = ArrayFromCollection
    If LBound(Tmpa) < UBound(Tmpa) Then 'single member trap
        For I = LBound(Tmpa) To UBound(Tmpa) - 1  '-1 means that a 2 member collection only gets shuffled once
            SwapAnyThing Tmpa(Int(Rnd * UBound(Tmpa))), Tmpa(Int(Rnd * UBound(Tmpa) + 1))
        Next I
    End If
    ArrayToCollection Tmpa

End Sub

Private Sub SwapAnyThing(element1 As Variant, element2 As Variant)

  'Swap any two items in any format

  Dim TEMP As Variant

    TEMP = element1
    element1 = element2
    element2 = TEMP

End Sub

Public Property Get Top() As Variant

  'Copyright 2002 Roger Gilchrist
  'Return largest member of collection
  'if numeric then numeric
  'if alphanumeric then lowest in alpha sort

    If Exists(m_Collection.Count, Silent) Then
        Top = m_Collection(m_Collection.Count)
    End If

End Property

':) Ulli's VB Code Formatter V2.13.6 (11/09/2002 10:26:21 AM) 76 + 509 = 585 Lines
