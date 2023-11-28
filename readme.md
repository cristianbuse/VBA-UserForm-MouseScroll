# VBA UserForm MouseScroll

MouseScroll is a VBA Project that allows Mouse Wheel Scrolling on MSForms Controls and Userforms.

- Multiple forms are tracked simultaneously. Just call the ```EnableMouseScroll``` for each form
- Both **MODAL** and **MODELESS** forms are supported (starting **12-Oct-2023**)!
- Debugging while mouse is hooked is now supported (starting **12-Oct-2023**)!
- This library can be extended for clicks, double-clicks and movement inputs
- Both **vertical** and **horizontal** scroll are supported. Hold down *Shift* key for horizontal scroll and *Ctrl* key for Zoom

## Installation

Just import the following 2 code modules in your VBA Project:

* [**MouseScroll.bas**](https://github.com/cristianbuse/VBA-UserForm-MouseScroll/blob/master/src/MouseScroll.bas)
* [**MouseOverControl.cls**](https://github.com/cristianbuse/VBA-UserForm-MouseScroll/blob/master/src/MouseOverControl.cls)

To avoid any issues with the ```CR``` and ```LF``` characters, it is best to download the available [ZIP](https://github.com/cristianbuse/VBA-UserForm-MouseScroll/archive/refs/heads/master.zip) and then import the modules from there.

## Usage
In your Modal Userform use:
```vba
EnableMouseScroll myUserForm
```
For example you can use your Form's Initialize Event:
```vba
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2

    EnableMouseScroll Me
End Sub
```

Not needed, but the following code can be added in the Form's Terminate Event:
```VBA
Private Sub UserForm_Terminate()
    DisableMouseScroll Me
End Sub
```
Tracking of forms is done automatically by checking if window is still valid and if the reference count of the form's object has any references left (except the internal ones used for raising events).

## Notes
* You can download the available Demo Workbook for a quick start

## Other Controls
* ```ListView```, ```TreeView```, ```WebBrowser``` etc. controls are supported without need for any changes to the code

## License
MIT License

Copyright (c) 2019 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
