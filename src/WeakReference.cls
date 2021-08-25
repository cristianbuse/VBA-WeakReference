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
'' In order to achieve a "Fake" Weak Reference in VBA, an external API that can
''    manipulate memory directly is needed (see CopyMemory below)
'' The easiest solution (not the one implemented in this class - see below) is
''    to store the target's object memory address and to create a New Reference
''    from the address whenever needed. Something like:
''       Dim ptr As LongPtr: ptr = ObjPtr(targetObj)
''       Dim size As LongPtr: size = LenB(ptr)
''    And later retrieve with:
''       Dim tempObject As Object
''       CopyMemory tempObj, ptr, size
''    The temporary object is not counted and would crash the Application
''       because the IUnknown::AddRef was not called (as it would when using the
''       usual assignment mechanism). So, after assigning the result, the
''       temporary object needs to be cleared using the same API:
''       Set Result = tempobj
''       CopyMemory tempobj, 0, size 'so reference is not decremented
''       'OR
''       ZeroMemory tempobj, size
''    Or, instead of clearing the uncounted reference, the reference count could
''    be manually incremented using something like (for VBA classes only):
''       Sub IncrementRefCount(iUnk As IUnknown, increment As Long)
''           Dim refCount As Long
''           Dim targetPtr As LongPtr: targetPtr = ObjPtr(iUnk) + LenB(refCount)
''           '
''           CopyMemory refCount, ByVal targetPtr, 4 'Gets current count
''           refCount = refCount + increment
''           CopyMemory ByVal targetPtr, refCount, 4 'Sets new count
''       End Sub
''       'With a call like:
''       IncrementRefCount tempObj, 1 'Note -1 would decrement the count
''    By incrementing the count, the tempObj is safely decremented later by VBA
''       so there is no need to clear the memory. But, costs an extra API call.
''
'' Using the above solution would require at least two calls to the
''    CopyMemory API everytime the Object is needed. This can be costly if the
''    method retrieving the Object is called many times (e.g. millions of times)
'' Instead, can this be achieved with plain VBA?
'' Yes!
''
'' ---------------------------------------
'' SOLUTION
'' ---------------------------------------
'' CopyMemory API is still needed but only once in the Class_Initalize event.
''
'' Looking at how a Variant is stored in memory:
''    'https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant?redirectedfrom=MSDN
''    one quickly realizes that the only difference between two Variants, one
''    holding an Object Reference and one holding an Integer (Long or LongLong),
''    is the vt (VarType) member (the first 2 bytes in any Variant).
'' Being able to manipulate the varType 2 bytes would allow to simply flip a
''    Variant between a Long Integer (memory address) to an Object and back
''    (i.e. setting vt to vbObject when the object is needed and then setting it
''    back to vbLong/vbLongLong to avoid any unwanted memory reclaims).
'' Changing var type using the same API would look like:
''    Dim ptr As Variant: ptr = ObjPtr(targetObj)
''    CopyMemory ptr, vbObject, 2
''    'ptr is now an Object - uncounted so still unsafe
''    Set Result = ptr 'Result is a counted reference
''    CopyMemory ptr, vbLong(or vbLongLong), 2
''    'ptr is now a Long/LongLong Integer - can safely go out of scope
'' Of course, changing 'vt' by using the same API is easy but that would defeat
''    the purpose of using plain VBA to make the retrival faster.
''
'' Solution is to use a redirected Variant of type Integer that remotely
''    manipulates the 2 bytes in the reference Variant. This can be done by
''    setting the VT_BYREF flag once, in the class initialize event. After
''    redirection, a simple VBA assignment would do the trick.
''    e.g. redirectedVariant = vbObject would change the reference data type
''    to Object and redirectedVariant = vbLong/vbLongLong would change it back
''
'' To make retrieval safe, the target object must implement the IWeakable
''    interface and call the 'ObjectTerminated' method from it's Class_Terminate
''    event for all weak references pointing to it. Or, it could delegate the
''    the responsibility to a separate contained object that would surely get
''    terminated at the same time (e.g. WeakRefInformer)
'*******************************************************************************
Option Explicit

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If

'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
'https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant?redirectedfrom=MSDN
'Flag used to simulate ByRef Variants
Private Const VT_BYREF As Long = &H4000

Private Type FAKE_REFERENCE
    remoteVarType As Variant  'Manipulates the variant type for 'reference'
    reference As Variant      'Will be holding the object reference/address
End Type

#If Win64 Then
    #If Mac Then
        Const vbLongLong As Long = 20 'Apparently missing for x64 on Mac
    #End If
    Const vbLongPtr As Long = vbLongLong
#Else
    Const vbLongPtr As Long = vbLong
#End If

Private m_fake As FAKE_REFERENCE

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class Constructor
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    'Save address of the Variant that will hold the target reference/address
    m_fake.remoteVarType = VarPtr(m_fake.reference)
    '
    'Change remoteVT variant type to Integer ByRef. This will now be linked
    '   to the first 2 bytes of the Variant holding the target reference
    'Setting the VT_BYREF flag makes sure that the 2 bytes are not reclaimed
    '   twice when both 'remoteVarType' and 'reference' go out of scope
    'And most importantly this gives the ability to switch the variant type of
    '   the reference at will, just by changing the Integer value of remoteVT
    CopyMemory m_fake.remoteVarType, vbInteger + VT_BYREF, 2
End Sub

'*******************************************************************************
'Saves the memory address to an object's default interface (not to IWeakable)
'*******************************************************************************
Public Property Set Object(obj As IWeakable)
    m_fake.reference = ObjPtr(GetDefaultInterface(obj))
    If m_fake.reference = 0 Then Exit Property
    '
    obj.AddWeakRef Me
End Property

'*******************************************************************************
'Returns the default interface for an object
'All VB intefaces are dual interfaces meaning all interfaces are derived from
'   IDispatch which in turn is derived from IUnknown. In VB the Object datatype
'   stands for the IDispatch interface.
'Casting from a custom interface (derived only from IUnknown) to IDispatch
'   forces a call to QueryInterface for the IDispatch interface (which knows
'   about the default interface)
'*******************************************************************************
Private Function GetDefaultInterface(obj As IUnknown) As Object
    Set GetDefaultInterface = obj
End Function

'*******************************************************************************
'Retrieves the object pointed by the saved reference
'No external API calls are needed!
'*******************************************************************************
Public Property Get Object() As Object
    If m_fake.reference = 0 Then Exit Property
    '
    Set Object = DeReferenceByVarType(m_fake.remoteVarType)
End Property

'*******************************************************************************
'Utility function needed to redirect remoteVT - See Class_Initialize comments
'*******************************************************************************
Private Function DeReferenceByVarType(ByRef remoteVT As Variant) As Object
    remoteVT = vbObject
    Set DeReferenceByVarType = m_fake.reference
    remoteVT = vbLongPtr
End Function

'*******************************************************************************
'Needs to be called when the referenced object is terminated
'*******************************************************************************
#If VBA7 Then
Public Sub ObjectTerminated(refAddress As LongPtr)
#Else
Public Sub ObjectTerminated(refAddress As Long)
#End If
    If m_fake.reference = refAddress Then m_fake.reference = 0
End Sub

'*******************************************************************************
'Returns the currently referenced memory address
'*******************************************************************************
#If VBA7 Then
Public Function ReferencedAddress() As LongPtr
#Else
Public Function ReferencedAddress() As Long
#End If
    ReferencedAddress = m_fake.reference
End Function
