VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSafeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2002 Roger Gilchrist
'This was just a spin-off from ClsStackCollection.
'I have not tested it fully and only include it for information purposes.
'It is a wrapper for VB collections with built in error detection and some auto-correction powers
'You might like to enhance it with some of the additional method and properties from ClsStackCollection
'or merge them into one class(just copy Add and Remove and some of the Msg in About; clsStackCollection already has the other routines)

'Search for '*---PROGRAMMER MODIFICATION POINT---*
'to see where you might like to modify the class for release purposes

Option Explicit
Private Verbose As Boolean
Private AutoSafety As Boolean
Private Const Silent As Boolean = True
Private m_Collection As New Collection

Public Sub About()

  'Copyright 2002 Roger Gilchrist
  'NOTE VB help comments "The maximum length of prompt is approximately 1024 characters,
  '                        depending on the width of the characters used."
  'This is why I have used vbCr rather than vbNewLine it saves one character per line

  Dim msg As String

    msg = "Just a work in progress. NOT used in StackCollection Demo." & vbCr & vbCr & _
          "Add safety nets and auto-corrections to VB's collections:" & vbCr & vbCr & _
          "ENHANCEMENTS" & vbCr & _
          "-----------------------------------------------------------" & vbCr & _
          "Add:" & vbTab & "Deals with the various error conditions that can occur" & vbCr & _
          "Count:" & vbTab & "VB standard Collection.Count" & vbCr & _
          "Item:" & vbTab & "Deals with the 'Index/Key does not exist' error" & vbCr & _
          "Remove:" & vbTab & "Deals with the 'Index/Key does not exist' error" & vbCr & vbCr & _
          "ADDITIONS" & vbCr & _
          "-----------------------------------------------------------" & vbCr & _
          "Exists IndexKey" & vbTab & vbTab & vbTab & "Test member exists" & vbCr & _
          "Replace IndexKey, NewValue" & vbTab & vbTab & "Edit a stack member" & vbCr & vbCr & _
          "Hope you like it!" & vbCr & _
          "Copyright 2002 Roger Gilchrist"
    MsgBox msg, vbInformation, "ClsCollection"

End Sub

Public Sub Add(Item As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)

  'This is a wrapper for the standard VB Add method with built in Error handlers
  'Copyright 2002 Roger Gilchrist
  
  Dim AddAction As Integer

    '*---PROGRAMMER MODIFICATION POINT---*
    AddAction = vbYesNo + vbDefaultButton2 + vbExclamation
    'AddAction defaults to No you might like to change that

    If Not IsMissing(Key) Then
        If Exists(Key) Then
            If Verbose Then
                If MsgBox("Would you like to replace the existing Item instead?", AddAction, "Item cannot be added") = vbYes Then
                    Replace Key, Item
                  Else 'NOT MSGBOX("WOULD YOU LIKE TO REPLACE THE EXISTING ITEM INSTEAD?",...
                    '*---PROGRAMMER MODIFICATION POINT---*
                    'add a way to modify the Key value here if you want
                    'otherwise add code to deal with failure externally to class
                End If
              ElseIf AutoSafety Then 'VERBOSE = FALSE
                Replace Key, Item
                '*---PROGRAMMER MODIFICATION POINT---*
                'OR add a way to modify the Key value here if you want
                'otherwise add code to deal with failure externally to class
            End If
            Exit Sub '>---> Bottom
        End If
    End If

    If Not IsMissing(Before) Then
        If Not Exists(Before) Then
            If Verbose Then
                If MsgBox("Would you like to proceed without Before: '" & Before & "'?", AddAction, "Item cannot be added") = vbYes Then
                    Set Before = Nothing
                    GoTo EndBeforeTest
                End If
              ElseIf AutoSafety Then 'VERBOSE = FALSE
                Set Before = Nothing
                GoTo EndBeforeTest
            End If
            '
            Exit Sub '>---> Bottom
        End If
    End If
EndBeforeTest:

    If Not IsMissing(After) Then
        If Not Exists(After) Then
            If Verbose Then
                If MsgBox("Would you like to proceed without After: '" & After & "'?", AddAction, "Item cannot be added") = vbYes Then
                    Set After = Nothing
                    GoTo EndAfterTest
                End If
              ElseIf AutoSafety Then 'VERBOSE = FALSE
                Set After = Nothing
                GoTo EndAfterTest
            End If
            '
            Exit Sub '>---> Bottom
        End If
    End If
EndAfterTest:

    m_Collection.Add Item, Key, Before, After

End Sub

Private Sub Class_Initialize()

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'Change Verbose to False 1. if you want to test the AutoSafety error handling
  '                        2. You probably don't want end-users to see the messages
  '
  'AutoSafety: Depending on how you want to handle errors you might want to set this to False
  '            if you don't like the classes default corrections or change the defaults

    Verbose = True     ' Use MsgBox to warn of error and offer a way out
    AutoSafety = Not Verbose ' Take the way out if it exists

    ' NOTE: If Verbose is True then AutoSafety will never trigger no matter what its value

End Sub

Public Property Get Count() As Variant

  'Wrapper for VB standard Collection.Count

    Count = m_Collection.Count

End Property

Public Property Get Exists(IndexKey As Variant, Optional Quiet As Boolean = False) As Boolean

  'Copyright 2002 Roger Gilchrist
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
            MsgBox msg, vbCritical, "ClsSafeCollection Error"
        End If
    End If

End Property

Public Property Get Item(IndexKey As Variant) As Variant

  'Wrapper for standard Item but with added safety returns an empty Variant on error
  'Copyright 2002 Roger Gilchrist

    If Exists(IndexKey) Then
        Item = m_Collection.Item(IndexKey)
      Else 'EXISTS(INDEXKEY) = FALSE
        If AutoSafety Then 'VERBOSE = FALSE
            ''*---PROGRAMMER MODIFICATION POINT---*
            ''there is no AutoSafety for this condition
            'you might like to supply one like this
            ''Item = "***MISSING ITEM(" & IndexKey & ")***"
            ''
            ''as is' an empty Variant is returned

        End If
    End If

End Property

Public Sub Remove(IndexKey As Variant)

  'Wrapper for VB standard Collection.Remove with added safety of no error if it does not exist
  'Copyright 2002 Roger Gilchrist
  ''*---PROGRAMMER MODIFICATION POINT---*
  ''This does not require an error message.
  ''If it doesn't exist that's what you wanted anyway
  ''but here it is in case you want it.
  ''Just uncomment the whole block once (the comments have two comment marks)
  '' you should also delete the Silent paramater

    If Exists(IndexKey, Silent) Then
        m_Collection.Remove IndexKey
        '      Else 'EXISTS(INDEXKEY) = FALSE
        '        If Verbose Then
        '            MsgBox "Your Index/Key: '" & IndexKey & "' does not exist.", vbOKOnly + vbInformation, "Item does not exist."
        '        ElseIf AutoSafety Then
        ''*---PROGRAMMER MODIFICATION POINT---*
        ''there is no AutoSafety for this condition
        ''you might like to supply one
        ' End If
    End If

End Sub

Public Sub Replace(IndexKey As Variant, NewValue As Variant)

  'Works by creating a temporary copy of m_collection but replacing the IndexKey value with your NewValue
  'because it has to work through the whole collection it will be slow for large collections
  'Copyright 2002 Roger Gilchrist
  
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

':) Ulli's VB Code Formatter V2.13.6 (11/09/2002 10:26:10 AM) 15 + 231 = 246 Lines
