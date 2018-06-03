Dim this_Sheet As Worksheet
Dim current_toggle_State As Boolean

Private Sub CommandButton1_Click()
    '7. When the user presses submit, take value from list, write to sheet
    this_Sheet.Range("D2").Value = Me.ComboBox1.Value
End Sub

Private Sub ToggleButton1_Change()
    '6. We then get the current state of the button again...
    current_toggle_State = Me.ToggleButton1.Value

    Select Case current_toggle_State
        '6. If I am still pressed, then...
        Case True:
            'Value of previous state is 1 (or pressed)
            this_Sheet.Range("D7").Value = 1
            'Change caption from LIST 1 to LIST 2
            ToggleButton1.Caption = "LIST 2"
            'List is still LIST 1...
            ComboBox1.Clear
            ComboBox1.AddItem "Tomato", 0
            ComboBox1.AddItem "Cucumber", 1
            ComboBox1.ListIndex = 0
        '6. If I am now unpressed by the user...
        Case False:
            'Value of previous state is 0 (or unpressed)
            this_Sheet.Range("D7").Value = 0
            'Change caption from LIST 2 to LIST 1
            ToggleButton1.Caption = "LIST 1"
            'LIST 2 to LIST 1...
            ComboBox1.Clear
            ComboBox1.AddItem "Pizza", 0
            ComboBox1.AddItem "Soda", 1
            ComboBox1.ListIndex = 0
    End Select
End Sub

Private Sub UserForm_Initialize()
    '1. We set our variables
    Dim previous_Toggle_State As Long
    
    '2. We set the sheet "food" as an object
    Set this_Sheet = Worksheets("Food")
    
    '3. We get the previous button state from the sheet...
    previous_Toggle_State = this_Sheet.Range("D7").Value
    
    Select Case previous_Toggle_State
        Case 1:
            '4. I was pressed, then button should be pressed...
            Me.ToggleButton1.Value = True
            'And, LIST 1 should active
            ComboBox1.Clear
            ComboBox1.AddItem "Tomato", 0
            ComboBox1.AddItem "Cucumber", 1
            ComboBox1.ListIndex = 0
            '4. Otherwise, it should be unpressed
        Case 0:
            Me.ToggleButton1.Value = False
            'LIST 2 should active
            ComboBox1.Clear
            ComboBox1.AddItem "Pizza", 0
            ComboBox1.AddItem "Soda", 1
            ComboBox1.ListIndex = 0
    End Select
End Sub
