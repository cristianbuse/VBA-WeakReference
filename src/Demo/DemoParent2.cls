VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoParent2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWeakable

Private m_children As Collection

Private Sub Class_Initialize()
    Set m_children = New Collection
End Sub

Public Sub AddChild(ch As DemoChild2)
    m_children.Add ch
    ch.Name = "Child " & m_children.Count
End Sub

Public Property Get Child(index_ As Long) As DemoChild2
    Set Child = m_children(index_)
End Property

Public Property Get ChildCount()
    ChildCount = m_children.Count
End Property

Private Sub Class_Terminate()
    Set m_children = Nothing
    Debug.Print "Parent terminated " & Now
End Sub

Private Sub IWeakable_AddWeakRef(wRef As WeakReference)
    Static informer As New WeakRefInformer
    informer.AddWeakRef wRef, Me
End Sub

