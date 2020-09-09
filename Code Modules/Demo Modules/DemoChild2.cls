VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoChild2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_parent As WeakReference
Private m_name As String

Public Property Set Parent(newParent As DemoParent2)
    Set m_parent = New WeakReference
    Set m_parent.Object = newParent
End Property
Public Property Get Parent() As DemoParent2
    Set Parent = m_parent.Object
End Property

Private Sub Class_Terminate()
    Debug.Print m_name & " terminated " & Now
End Sub

Public Property Let Name(newName As String)
    m_name = newName
End Property
Public Property Get Name() As String
    Name = m_name
End Property
