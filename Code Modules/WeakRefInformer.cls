VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeakRefInformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA WeakReference
'''--------------------------------------------------
''' https://github.com/cristianbuse/VBA-WeakReference
'''--------------------------------------------------
'''
''' Copyright (c) 2020 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to deal
''' in the Software without restriction, including without limitation the rights
''' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''' copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in all
''' copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''' SOFTWARE.
'''=============================================================================

'*******************************************************************************
'' When terminated, informs all stored WeakReference objects about termination
'*******************************************************************************

Option Explicit

Private m_refs As Collection
#If VBA7 Then
    Private m_reference As LongPtr
#Else
    Private m_reference As Long
#End If

Public Sub AddWeakRef(wRef As WeakReference, obj As IWeakable)
    'Store the address for the object implementing IWeakable
    'When Class_Terminate is triggered, this will be passed to each
    '   WeakReference object just in case the WeakReference will be set to a
    '   different target (in-between this call and the termination call)
    If m_reference = 0 Then m_reference = ObjPtr(GetDefaultInterface(obj))
    '
    If wRef.ReferencedAddress = m_reference Then m_refs.Add wRef
End Sub

Private Function GetDefaultInterface(obj As IUnknown) As Object
    Set GetDefaultInterface = obj
End Function

Private Sub Class_Initialize()
    Set m_refs = New Collection
End Sub

Private Sub Class_Terminate()
    Dim wRef As WeakReference
    '
    For Each wRef In m_refs
        wRef.ObjectTerminated m_reference
    Next wRef
    Set m_refs = Nothing
End Sub
