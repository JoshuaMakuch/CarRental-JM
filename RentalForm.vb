'Joshua Makuch
'RCET 0265
'Spring 2023
'Car Rental Form Fork
'https://github.com/JoshuaMakuch/CarRental-JM.git

'Things to do:
'Code for just one customer and accumulate all the necessary data, validate, and then display

Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm

    'Asks the user if they want to exit the program and does an appropriate reaction
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If MessageBox.Show("Do you want to exit?", "Exit Program", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    'Asks the user if they want to exit the program and does an appropriate reaction but for the Tool Strip Menu
    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        ExitButton_Click(sender, e)
    End Sub

    'What happens when the user clicks the clear button
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        'Assures the user wants to clear all data
        If MessageBox.Show("Do you want to clear all data?", "Clear Data?", MessageBoxButtons.YesNo) = DialogResult.Yes Then

        End If

    End Sub

    'What happens when the user presses the calculate button
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        'Adds all input text boxs to an array to work with
        Dim inputTextBoxes() As TextBox = {NameTextBox, AddressTextBox, CityTextBox, StateTextBox, ZipCodeTextBox,
            BeginOdometerTextBox, EndOdometerTextBox, DaysTextBox}
        'What to display in a message to the user
        Dim messageInfo As String = ""
        'Creates a new variable to work with to reduce using the same "inputTextBoxes(i)" over and over again and for validation
        Dim tb As TextBox
        Dim validInput As Boolean = True

        'Checks the user has input at least something in all fields from the end to the beginning tab order
        For i As Integer = inputTextBoxes.Count - 1 To 0 Step -1

            'Resets each variable
            tb = inputTextBoxes(i)
            validInput = True

            'Checks to make sure there is a valid value in the user's input for that specific text box, if not, display and hightlight
            If tb.Text.Trim().Length = 0 Then
                validInput = False
            ElseIf tb Is BeginOdometerTextBox Or tb Is EndOdometerTextBox Then
                Try
                    If CInt(BeginOdometerTextBox.Text) > CInt(EndOdometerTextBox.Text) Then
                        validInput = False
                    End If
                Catch ex As Exception
                    validInput = False
                End Try
            ElseIf tb Is DaysTextBox Then
                Try
                    If CInt(tb.Text) <= 0 Or CInt(tb.Text) > 45 Then
                        validInput = False
                    End If
                Catch ex As Exception
                    validInput = False
                End Try
            End If

            'Checks the valid input variable and executes the appropriate response
            If validInput Then
                tb.BackColor = Color.LightGreen
            Else
                messageInfo = $"{messageInfo}{vbCrLf}{Chr(34)}{tb.Text}{Chr(34)} was an invalid value for {tb.Name}."
                tb.BackColor = Color.LightPink
                tb.Focus()
                tb.Text = ""
            End If

        Next

        'Displays a message box if the input data is wrong
        If messageInfo IsNot "" Then
            MessageBox.Show($"Identified Issues:{vbCrLf}{messageInfo}")
        End If
    End Sub

    'What happens when the user presses the menu strip calculate button
    Private Sub CalculateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateToolStripMenuItem.Click
        'Runs the calculate function when the menu strip item is clicked
        CalculateButton_Click(sender, e)
    End Sub

End Class
