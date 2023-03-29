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

    'What happens when the user clicks the clear button
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        'Assures the user wants to clear all data
        If MessageBox.Show("Do you want to clear all data?", "Clear Data?", MessageBoxButtons.YesNo) = DialogResult.Yes Then

        End If

    End Sub

    'What happens when the user presses the calculate button
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        'Runs the calculate sub when the calculate button is clicked
        Calculate()
    End Sub
    'What happens when the user presses the menu strip calculate button
    Private Sub CalculateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateToolStripMenuItem.Click
        'Runs the calculate function when the menu strip item is clicked
        Calculate()
    End Sub

    'What happens when the user submits their information and also validates that the inputs are correct
    Sub Calculate()
        'Adds all input text boxs to an array to work with
        Dim inputTextBoxes() As TextBox = {NameTextBox, AddressTextBox, CityTextBox, StateTextBox, ZipCodeTextBox,
            BeginOdometerTextBox, EndOdometerTextBox, DaysTextBox}
        'What to display in a message to the user
        Dim messageInfo As String = ""

        'Checks the user has input at least something in all fields from the end to the beginning tab order
        For i As Integer = inputTextBoxes.Count - 1 To 0 Step -1

            'Assumes the input is valid
            inputTextBoxes(i).BackColor = Color.LightGreen

            'Checks to make sure there is something in the user's input for that specific text box, if not, display and hightlight
            If inputTextBoxes(i).Text.Trim().Length = 0 Then
                messageInfo = $"{messageInfo}{vbCrLf}{inputTextBoxes(i).Name} is empty."
                inputTextBoxes(i).BackColor = Color.LightPink
                inputTextBoxes(i).Focus()
            ElseIf inputTextBoxes(i) Is EndOdometerTextBox Then
                Try
                    If CDbl(EndOdometerTextBox.Text) > CDbl(BeginOdometerTextBox.Text) Then
                        messageInfo = $"{messageInfo}{vbCrLf}{inputTextBoxes(i).Name} is not a possible amount."
                        inputTextBoxes(i).BackColor = Color.LightPink
                        inputTextBoxes(i).Focus()
                    End If
                Catch ex As Exception
                    messageInfo = $"{messageInfo}{vbCrLf}{inputTextBoxes(i).Name} is an invalid value."
                    inputTextBoxes(i).BackColor = Color.LightPink
                    inputTextBoxes(i).Focus()
                End Try
            ElseIf inputTextBoxes(i) Is DaysTextBox Then
                Try
                    If CInt(inputTextBoxes(i).Text) <= 0 Or CInt(inputTextBoxes(i).Text) > 45 Then
                        messageInfo = $"{messageInfo}{vbCrLf}{inputTextBoxes(i).Name} is an invalid amount."
                        inputTextBoxes(i).BackColor = Color.LightPink
                        inputTextBoxes(i).Focus()
                    End If
                Catch ex As Exception
                    messageInfo = $"{messageInfo}{vbCrLf}{inputTextBoxes(i).Name} is an invalid value."
                    inputTextBoxes(i).BackColor = Color.LightPink
                    inputTextBoxes(i).Focus()
                End Try
            End If

        Next

        'Displays a message box if the input data is wrong
        If messageInfo IsNot "" Then
            MessageBox.Show($"Identified Issues:{vbCrLf}{messageInfo}")
        End If

    End Sub

End Class
