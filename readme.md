# VBA UserForm MouseScroll

MouseScroll is a VBA Project that allows Mouse Wheel Scrolling on MSForms Controls and Userforms but can also be extended for clicks, double-clicks and movement inputs.

## Installation

Just import the following 2 code modules in your VBA Project:

* **MouseScroll.bas**  
* **MouseOverControl.cls**

## Usage
In your Modal Userform use:
```vba
HookMouseToForm Me
```
For example you can use your Form's Initialize Event:
```vba
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2

    HookMouseToForm Me
End Sub
```

## Notes
* Hold Shift for Horizontal Scroll and Ctrl for Zoom
* The Mouse Hook will not work with Modeless Forms (Modal only)
* No need to call the Unhook method. It will be called automatically when the Form is inactive
* If you call a second Modal Form make sure to Hook back the first one when done:
```vba
Private Sub ShowSecondForm_Click()
    SecondForm.Show vbModal
    MouseScroll.HookMouseToForm Me
End Sub
```
* You can download the available Demo Workbook for a quick start

## License
MIT License

Copyright (c) 2019 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.