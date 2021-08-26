VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_parent As WeakReference

Public Property Set Parent(ByVal newParent As DemoParent)
    Set m_parent = New WeakReference
    Set m_parent.Object = newParent
End Property
Public Property Get Parent() As DemoParent
    Set Parent = m_parent.Object
End Property

Private Sub Class_Terminate()
    Debug.Print "Child terminated " & Now
End Sub
