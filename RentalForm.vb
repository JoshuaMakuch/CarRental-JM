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
        'Adds all customer data to an array
        Dim inputTextBoxes() As TextBox = {NameTextBox, AddressTextBox, CityTextBox, StateTextBox, ZipCodeTextBox,
            BeginOdometerTextBox, EndOdometerTextBox, DaysTextBox}

        'Checks the user has input at least something into all fields.
        For Each tb As TextBox In inputTextBoxes
            If tb.Text.Trim().Length = 0 Then
                MessageBox.Show("You must enter a value for all fields.")
                tb.Focus()
                Return
            End If
            If tb Is DaysTextBox Then
                Try
                    If CInt(DaysTextBox.Text) <= 0 Or CInt(DaysTextBox.Text) > 45 Then
                        MessageBox.Show("Please input a valid amount of days (0 to 45 days).")
                        tb.Focus()
                        Return
                    End If
                Catch ex As Exception
                    MessageBox.Show("Please input a valid amount of days (0 to 45 days).")
                End Try
            End If
        Next

    End Sub

End Class
