Attribute VB_Name = "DemoWeakRef"
Option Explicit

Const LOOPS As Long = 1000000

Sub DemoMain()
    DemoTerminateParentFirst
    DemoTerminateChildFirst
    DemoTerminateParentBeforeClosing
    DemoSetParentAgainBeforeClosing
    DemoTerminateExecutionWithExtraChildRef 'stops execution
    '
    'Code will not reach this line so you need to run this method manually
    DemoTerminateExecutionWithNoExtraRefs
End Sub

'>>>>>>>>>>>>>>>>>>
'Demo Methods below
'>>>>>>>>>>>>>>>>>>

Sub DemoTerminateParentFirst()
    Debug.Print String$(62, "=")
    Debug.Print "DemoTerminateParentFirst"
    Debug.Print "..."
    '
    Dim c As New DemoChild
    Dim p As New DemoParent
    '
    p.Child = c
    c.Parent = p
    '
    Dim t As Double
    Dim i As Long
    Dim s As String
    '
    t = Timer
    For i = 1 To LOOPS
        s = TypeName(p.Child.Parent)
        If s <> "DemoParent" Then Stop
    Next i
    Debug.Print "Retrieved Parent from Child for " & LOOPS _
        & " times in " & Round(Timer - t, 3) & " seconds"
End Sub

Sub DemoTerminateChildFirst()
    Debug.Print String$(62, "=")
    Debug.Print "DemoTerminateChildFirst"
    Debug.Print "..."
    '
    Dim c As New DemoChild
    Dim p As New DemoParent
    '
    p.Child = c
    c.Parent = p
    Set c = Nothing
    '
    Dim t As Double
    Dim i As Long
    Dim s As String
    '
    t = Timer
    For i = 1 To LOOPS
        s = TypeName(p.Child.Parent)
        If s <> "DemoParent" Then Stop
    Next i
    Debug.Print "Retrieved Parent from Child for " & LOOPS _
        & " times in " & Round(Timer - t, 3) & " seconds"
End Sub

Sub DemoTerminateParentBeforeClosing()
    Debug.Print String$(62, "=")
    Debug.Print "DemoTerminateParentBeforeClosing"
    Debug.Print "..."
    '
    Dim c As New DemoChild
    Dim p As New DemoParent
    '
    p.Child = c
    c.Parent = p
    '
    Dim t As Double
    Dim i As Long
    Dim s As String
    '
    t = Timer
    For i = 1 To LOOPS
        s = TypeName(p.Child.Parent)
        If s <> "DemoParent" Then Stop
    Next i
    Debug.Print "Retrieved Parent from Child for " & LOOPS _
        & " times in " & Round(Timer - t, 3) & " seconds"
    '
    Set p = Nothing
    Debug.Print "Parent is now: " & TypeName(c.Parent)
End Sub

Sub DemoSetParentAgainBeforeClosing()
    Debug.Print String$(62, "=")
    Debug.Print "DemoSetParentAgainBeforeClosing"
    Debug.Print "..."
    '
    Dim c As New DemoChild
    Dim p As New DemoParent
    '
    p.Child = c
    c.Parent = p
    '
    Dim t As Double
    Dim i As Long
    Dim s As String
    '
    t = Timer
    For i = 1 To LOOPS
        s = TypeName(p.Child.Parent)
        If s <> "DemoParent" Then Stop
    Next i
    Debug.Print "Retrieved Parent from Child for " & LOOPS _
        & " times in " & Round(Timer - t, 3) & " seconds"
    '
    c.Parent = Nothing
    c.Parent = p
    Debug.Print "Parent is now: " & TypeName(c.Parent)
End Sub

'The End statement stops code execution abruptly, without invoking the
'Unload, QueryUnload, or Terminate event, or any other Visual Basic
'code. Code you have placed in the Unload, QueryUnload, and Terminate
'events of forms and class modules is not executed. Objects created
'from class modules are destroyed, files opened by using the Open
'statement are closed, and memory used by your program is freed.
'Object references held by other programs are invalidated.
'Quoted from:
'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/end-statement

Sub DemoTerminateExecutionWithExtraChildRef()
    Debug.Print String$(62, "=")
    Debug.Print "DemoTerminateExecutionWithExtraChildRef"
    Debug.Print "..."
    '
    Dim c As New DemoChild
    Dim p As New DemoParent
    '
    p.Child = c
    c.Parent = p
    '
    Dim t As Double
    Dim i As Long
    Dim s As String
    '
    t = Timer
    For i = 1 To LOOPS
        s = TypeName(p.Child.Parent)
        If s <> "DemoParent" Then Stop
    Next i
    Debug.Print "Retrieved Parent from Child for " & LOOPS _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Print "Stopping Execution"
    End 'This does not cause any crashes because the Weak Ref is safe
End Sub

Sub DemoTerminateExecutionWithNoExtraRefs()
    Debug.Print String$(62, "=")
    Debug.Print "DemoTerminateExecutionWithNoExtraRefs"
    Debug.Print "..."
    '
    Dim c As New DemoChild
    Dim p As New DemoParent
    '
    p.Child = c
    c.Parent = p
    Set c = Nothing
    '
    Dim t As Double
    Dim i As Long
    Dim s As String
    '
    t = Timer
    For i = 1 To LOOPS
        s = TypeName(p.Child.Parent)
        If s <> "DemoParent" Then Stop
    Next i
    Debug.Print "Retrieved Parent from Child for " & LOOPS _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Print "Stopping Execution"
    End 'This does not cause any crashes because the Weak Ref is safe
End Sub
