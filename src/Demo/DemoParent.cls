VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWeakable

Private m_child As DemoChild

Public Property Set Child(ByVal ch As DemoChild)
    Set m_child = ch
End Property
Public Property Get Child() As DemoChild
    Set Child = m_child
End Property

Private Sub Class_Terminate()
    Set m_child = Nothing
    Debug.Print "Parent terminated " & Now
End Sub

Private Sub IWeakable_AddWeakRef(ByVal wRef As WeakReference)
    Static informer As New WeakRefInformer
    informer.AddWeakRef wRef, Me
End Sub
