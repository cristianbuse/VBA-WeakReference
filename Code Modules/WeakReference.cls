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
''    2) Use a Weak Reference. A Weak Reference is not counted
''       (i.e. IUnknown::AddRef method is not called) and does not protect the
''       referenced object from being collected by the Garbage Collector (GC)
''       unlike a Strong Reference.
''       e.g. A->B and B--A
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
''       ZerMemory tempobj, size
''    Or, instead of clearing the uncounted reference, the reference count could
''    be manually incremented using something like:
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
'' CopyMemory API is still needed but only once in the Class_Initalize event and
''    once for each time the reference object is set (through the SetObject
''    method) which is likely to be only one time. So 2 API calls in total.
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
''    CopyMemory ByVal VarPtr(ptr), vbObject, 2
''    'ptr is now an Object - uncounted so still unsafe
''    Set Result = ptr 'Result is a counted reference
''    CopyMemory ByVal VarPtr(ptr), vbLong(or vbLongLong), 2
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
'' The added bonus is that a Variant that has the ByRef flag set, can safely go
''    out of scope as no memory is reclaimed
'' To make retrieval safer, a second redirected Variant can be used to point at
''    the first bytes of the object's address (holding virtual table pointer for
''    the default interface) and then comparing with the inital value of the
''    vTable pointer thus avoiding to turn the Variant into an Object type if
''    the target object has been destroyed already
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
'Flag used to simulate ByRef Variants in order to avoid memory reclaim
Private Const VT_BYREF As Long = &H4000 'Makes it all possible

'A memory address Long Integer
Private Type MEM_ADDRESS
    #If VBA7 Then
        ptr As LongPtr 'Defaults to LongLong on x64 or Long on x32
    #Else
        ptr As Long    'For VB6
    #End If
End Type

Private Type FAKE_REFERENCE
    remoteVarType As Variant 'Manipulates the variant type for 'reference'
    reference As Variant     'Will be holding the object reference/address
    vTable As MEM_ADDRESS    'Initial address of virtual table
    vTableByRef As Variant   'Address used to check if reference is still valid
    vbLongPtr As Long        'Data type enum (vbLong = 3 or vbLongLong = 20)
    isValid As Boolean       'Indicates if the memory reference is valid
End Type

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
    CopyMemory ByVal VarPtr(m_fake.remoteVarType), vbInteger + VT_BYREF, 2
    '
    'Store the data type enum for mem addresses (vbLong = 3 or vbLongLong = 20)
    m_fake.vbLongPtr = VBA.VarType(ObjPtr(Nothing))
End Sub

'*******************************************************************************
'Sets the weak/fake reference to an object
'*******************************************************************************
Public Property Let Object(obj As Object)
    'Save memory address of the object
    m_fake.reference = ObjPtr(obj)
    '
    m_fake.isValid = (m_fake.reference <> 0)
    If Not m_fake.isValid Then Exit Property
    '
    'Save the default interface's virtual table address by reference. The vTable
    '   address is found at the first 4 (x32) or 8 (x64) bytes at the referenced
    '   interface address
    m_fake.vTableByRef = m_fake.reference
    CopyMemory ByVal VarPtr(m_fake.vTableByRef), m_fake.vbLongPtr + VT_BYREF, 2
    '
    'Save the current vTable address. This is needed later to compare with the
    '   vTableByRef address in order to establish if the Object has been
    '   destroyed and it's memory reclaimed.
    'vTableByRef can still be read within the scope of this method
    m_fake.vTable.ptr = m_fake.vTableByRef
End Property

'*******************************************************************************
'Safely retrieves the object that the saved reference is pointing to
'No external API calls are needed!
'*******************************************************************************
Public Property Get Object() As Object
    If Not m_fake.isValid Then Exit Property
    '
    'Compare the current vTable address value with the initial address
    'The current redirected value vTableByRef can NOT be read directly anymore
    '   so it must be passed ByRef to an utility function
    m_fake.isValid = (GetRemoteAddress(m_fake.vTableByRef).ptr = m_fake.vTable.ptr)
    '
    If m_fake.isValid Then
        'Address is still valid. Retrive the object
        'Turn the reference into an object (needs to be done ByRef)
        VarType(m_fake.remoteVarType) = vbObject
        Set Object = m_fake.reference
    End If
    '
    'The fake object is not counted (reference count was never incremented by
    '   calling the IUnknown::AddRef method) so a crash will occur if the
    '   Variant type remains as vbObject, because when the Variant goes out
    '   of scope the object count is decremented one more time than it should
    'Meanwhile, as Integer, the Variant can safely go out of scope anytime
    VarType(m_fake.remoteVarType) = m_fake.vbLongPtr 'vbLong or vbLongLong
End Property

'*******************************************************************************
'Utility. Changes the data type for the reference Variant while preserving the
'   level of redirection of remoteVarType
'*******************************************************************************
Private Property Let VarType(ByRef v As Variant, newType As Integer)
    v = newType
End Property

'*******************************************************************************
'Returns the value of a Variant that has the VT_BYREF flag set
'*******************************************************************************
Private Function GetRemoteAddress(ByRef memAddress As Variant) As MEM_ADDRESS
    GetRemoteAddress.ptr = memAddress
End Function
