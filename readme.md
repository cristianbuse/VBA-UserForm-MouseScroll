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
Copyright (C) 2019 VBA Mouse Scroll project contributors

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see [http://www.gnu.org/licenses/](http://www.gnu.org/licenses/) or
[GPLv3](https://choosealicense.com/licenses/gpl-3.0/).