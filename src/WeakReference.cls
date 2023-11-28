VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA WeakReference
''' -------------------------------------------------
''' https://github.com/cristianbuse/VBA-WeakReference
''' -------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2020 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

'*******************************************************************************
'' ---------------------------------------
'' INTRO
'' ---------------------------------------
'' The infamous Reference Cycle problem means that objects are referencing each
''    other and will remain loaded in memory (not reclaimed) even after the
''    program has stopped execution.
''    e.g. A->B (object A holds a reference to object B) and viceversa B->A.
'' Solutions:
''    1) One of the references needs to be manually destroyed
''       e.g. Destroying B->A leaves only A->B . When A goes out of scope
''       then B goes out of scope (if no other references are pointing to B,
''       except A->B). Call would be like: B.TerminateRefs
''       Can be easy to forget to manually call the procedure
''    2) Use a Weak Reference. e.g. A->B and B--A. In a Garbage Collected
''       language, a Weak Reference is not counted and does not protect the
''       referenced object from being collected by the Garbage Collector (GC)
''       unlike a Strong Reference. In a non-GC language the weak references are
''       counted separately from the strong ones (e.g. SWIFT language)
'' Weak references are not a language feature in VBA. All references are Strong.
''
'' ---------------------------------------
'' SOLUTION
'' ---------------------------------------
'' A pointer is stored for the default interface of an object and the Object
''    is later retrieved by dereferencing the stored pointer.
''
'' To make retrieval safe, the target object must implement the IWeakable
''    interface and call the 'ObjectTerminated' method from it's Class_Terminate
''    event for all weak references pointing to it. Or, it could delegate the
''    the responsibility to a separate contained object that would surely get
''    terminated at the same time (e.g. WeakRefInformer)
''
'' Traditionally, an Object was dereferenced by using the CopyMemoryAPI but
''    that is extremely slow on certain systems. To avoid this issue, this
''    class uses 'LibMemory' (https://github.com/cristianbuse/VBA-MemoryTools)
''    to quickly read/write required memory
'*******************************************************************************
Option Explicit

Private m_ptr As LongPtr

'*******************************************************************************
'Saves the memory address of an object's default interface (not IWeakable)
'*******************************************************************************
Public Property Set Object(ByVal obj As IWeakable)
    m_ptr = ObjPtr(GetDefaultInterface(obj))
    If m_ptr = NULL_PTR Then Exit Property
    obj.AddWeakRef Me
End Property

'*******************************************************************************
'Retrieves the object pointed by the saved reference
'*******************************************************************************
Public Property Get Object() As Object
    If m_ptr = NULL_PTR Then Exit Property
    Set Object = MemObj(m_ptr)
End Property

'*******************************************************************************
'Needs to be called when the referenced object is terminated
'*******************************************************************************
Public Sub ObjectTerminated(ByVal refAddress As LongPtr)
    If m_ptr = refAddress Then m_ptr = 0
End Sub

'*******************************************************************************
'Returns the currently referenced memory address
'*******************************************************************************
Public Function ReferencedAddress() As LongPtr
    ReferencedAddress = m_ptr
End Function
