# VBA-WeakReference
Break Reference Cycles in VBA using simulated Weak Object References. In a Garbage Collected language, a Weak Reference is not counted and does not protect the referenced object from being collected by the Garbage Collector (GC) unlike a Strong Reference. In a non-GC language the weak references are counted separately from the strong ones (e.g. SWIFT language).

## Implementation
In order to be referenced by **WeakReference**, a VBA class must implement the **IWeakable** interface. A **WeakReference** instance saves the memory address of the targeted object (the address of the default interface - not the IWeakable interface) and later uses it to retrieve the target object.
A Variant is used for storing the target memory address. A second Variant is used to remotely manipulate the VarType of the first Variant by changing the first 2 bytes in memory.
When the object is needed, the first Variant is turned into an Object (unmanaged - not counted i.e. IUnknown::AddRef method is not called) by changing it's VarType to vbObject (remotely from the second Variant) and then the object is returned. The first variant is then changed back to an Integer (Long/LongLong) memory address to avoid decrementing the reference count of the first variant when it goes out of scope (which would crash the Application).

The remote manipulation of the VarType is done by setting the *VT_BYREF* flag on the second Variant. This is done by using a *CopyMemory* API but only once in the *Class_Initialize* event. After the flag is set, the second Variant can be used to change the VarType of the first Variant just by using a VBA assignment operation (needs a utility method for redirection though). No matter how many times the referenced object needs to be retrieved, no API calls are needed.

To be safe, all WeakReferences must be informed that the targeted Object has been terminated. Unfortunately, the Class_Terminate event is not part of the interface so it cannot be forced to do anything. Because too much boilerplate code would need to be added to all classes implementing IWeakable it is probably best to encapsulate all the logic inside a separate class called **WeakRefInformer** which is to be contained by the targeted class. The main idea is that by not exposing the contained WeakRefInformer object, it will surely go out of scope when the object implementing IWeakable is terminated.

A quick visual example. Consider a "parent" object containing 2 "child" objects pointing back through weak references and a 3rd "loose" weak reference. This would look like:  
![enter image description here](https://i.stack.imgur.com/7VhWj.png)

See full implementation description in the **WeakReference.cls** class

## Installation
Just import the following code modules in your VBA Project:
* **WeakReference.cls**
* **IWeakable.cls**
* **WeakRefInformer.cls**

## Usage
In all classes that need to be compatible with/referenced by ```WeakReference```, add the following code:
```VBA
Implements IWeakable

Private Sub IWeakable_AddWeakRef(wRef As WeakReference)
    Static informer As New WeakRefInformer
    informer.AddWeakRef wRef, Me
End Sub
```
Create a new instance of the **WeakReference** class and assign the desired object:
```VBA
Set wRef = New WeakReference
Set wRef.Object = targetObj 'targetObject implements IWeakable
```

Retrieving the object can be done at any time using:
```vba
Set obj = wRef.Object
```

## Demo

```Class1```:
```VBA
'Class1
Option Explicit

Implements IWeakable

Public x As Long

Private Sub IWeakable_AddWeakRef(wRef As WeakReference)
    Static informer As New WeakRefInformer
    informer.AddWeakRef wRef, Me
End Sub
```
Method in a regular code module:
```VBA
Sub TestWeakReference()
    Dim c As Class1
    Dim w1 As New WeakReference
    Dim w2 As New WeakReference
    Dim w3 As New WeakReference
    '
    Set c = New Class1
    c.x = 1
    '
    Set w1.Object = c
    Set w2.Object = c
    Set w3.Object = c
    
    Debug.Print w1.Object.x         'Prints 1 (correct)
    Debug.Print w2.Object.x         'Prints 1 (correct)
    Debug.Print w3.Object.x         'Prints 1 (correct)
    Debug.Print TypeName(w1.Object) 'Prints Class1 (correct)
    Debug.Print TypeName(w2.Object) 'Prints Class1 (correct)
    Debug.Print TypeName(w3.Object) 'Prints Class1 (correct)
    '
    Dim temp As Class1
    Set temp = New Class1
    Set w3.Object = temp
    temp.x = 2
    '
    Set c = Nothing 'Note this only resets w1 and w2 (not w3)
    Set c = New Class1
    c.x = 3
    '
    Debug.Print TypeName(w1.Object) 'Prints Nothing (correct)
    Debug.Print TypeName(w2.Object) 'Prints Nothing (correct)
    Debug.Print TypeName(w3.Object) 'Prints Class1 (correct)
    Debug.Print w3.Object.x         'Prints 2 (correct)
End Sub
```

## Testing

Import the following code modules:
* **DemoChild.cls**
* **DemoParent.cls**
* **DemoWeakRef.bas**
* **DemoChild2.cls**
* **DemoParent2.cls**

and execute method:
```vba
DemoWeakRef.DemoMain
```

## Notes
* There are no memory leaks even if state is lost.
* If the saved object has been destroyed, the WeakReference.Object (Get property) safely  returns Nothing
* The **WeakRefInformer.cls** is not really needed but avoids the duplication of the same code across all classes implementing IWeakable. Just the minimal code presented above in the **Usage** section is needed when using the informer.

## External contributions (not Git)
Many thanks to Matthieu ([GitHub](https://github.com/retailcoder) / [CR](https://codereview.stackexchange.com/users/23788/mathieu-guindon))  and Greedo ([GitHub](https://github.com/Greedquest) / [CR](https://codereview.stackexchange.com/users/146810/greedo)). See their contributions on [CodeReview](https://codereview.stackexchange.com/questions/245660/simulated-weakreference-class).

## License
MIT License

Copyright (c) 2020 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.