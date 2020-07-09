# VBA-WeakReference
Break Reference Cycles in VBA using simulated Weak Object References. A Weak Reference is not counted (i.e. IUnknown::AddRef method is not called) and does not protect the referenced object from being collected by the Garbage Collector, unlike a Strong Reference.

## Implementation
A Variant is used for storing an Object memory address (retrieved with ObjPtr(targetObj)). A second Variant is used to remotely manipulate the VarType of the first Variant by changing the first 2 bytes (of the first Variant) in memory.
When the object is needed, the first Variant is turned into an Object by changing it's VarType to vbObject (remotely from the second Variant) and then the object is returned. The first variant is then changed back to an Integer (Long/LongLong) memory address to avoid unwanted memory reclaim.

The remote manipulation of the VarType is done by setting the VT_BYREF flag on the second Variant. This is done by using a CopyMemory API but only once. After the flag is set, the second Variant can be used to change the VarType of the first Variant just by using a VBA assignment operation. No matter how many times the referenced object needs to be retrieved, no API calls are needed.

See full implementation description in the **WeakReference.cls** class

## Installation
Just import the following code module in your VBA Project:
* **WeakReference.cls**

## Usage
Create a new instance of the **WeakReference** class and assign the desired object
```vba
Set wRef = New WeakReference
wRef.SetObject targetObj
```

Retrieving the object can be done at any time using:
```vba
wRef.GetObject
```

## Testing

Import the following code modules:
* **DemoChild.cls**
* **DemoParent.cls**
* **DemoWeakRef.bas**

and execute method:
```vba
DemoWeakRef.DemoMain
```

## Notes
* There are no memory leaks even if state is lost.
* If the saved object has been destroyed, the .GetObject method safely  returns Nothing

## License
MIT License

Copyright (c) 2020 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.