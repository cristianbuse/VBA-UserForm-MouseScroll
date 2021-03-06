# VBA UserForm MouseScroll

MouseScroll is a VBA Project that allows Mouse Wheel Scrolling on MSForms Controls and Userforms but can also be extended for clicks, double-clicks and movement inputs.

Multiple forms are tracked simultaneously. Just call the ```EnableMouseScroll``` for each form.

## Installation

Just import the following 2 code modules in your VBA Project:

* [**MouseScroll.bas**](https://github.com/cristianbuse/VBA-UserForm-MouseScroll/blob/master/Code%20Modules/MouseScroll.bas)
* [**MouseOverControl.cls**](https://github.com/cristianbuse/VBA-UserForm-MouseScroll/blob/master/Code%20Modules/MouseOverControl.cls)

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

## Notes
* Hold Shift for Horizontal Scroll and Ctrl for Zoom
* The Mouse Hook will not work with Modeless Forms (Modal only)
* No need to call the ```DisableMouseScroll``` method. It will be called automatically when the Form's Window is destroyed
* Multiple forms are now tracked simultaneously and the mouse is unhooked automatically only when no forms are being tracked
* You can download the available Demo Workbook for a quick start

## License
MIT License

Copyright (c) 2019 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
