Option Explicit On
Option Strict On
Option Compare Binary
'Joshua Makuch
'RCET 0265
'Spring 2023
'Car Rental Form Fork
'https://github.com/JoshuaMakuch/CarRental-JM.git

'Things to do:
'Code for just one customer and accumulate all the necessary data, validate, and then display

Public Class RentalForm
    Dim sumCustomers As Integer
    Dim sumMilesDriven As Integer
    Dim sumCharges As Double

    'Asks the user if they want to exit the program and does an appropriate reaction
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If MessageBox.Show("Do you want to exit?", "Exit Program", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    'Menu Strip exit button
    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        ExitButton_Click(sender, e)
    End Sub

    'Clear button
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        'Assures the user wants to clear all data
        If MessageBox.Show("Do you want to clear all but summary data?", "Clear Data?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            'Adds all input text boxs to an array to work with
            Dim inputTextBoxes() As TextBox = {NameTextBox, AddressTextBox, CityTextBox, StateTextBox, ZipCodeTextBox,
            BeginOdometerTextBox, EndOdometerTextBox, DaysTextBox}
            Dim tb As TextBox
            For i As Integer = inputTextBoxes.Count - 1 To 0 Step -1
                'Resets appropriate textboxes
                tb = inputTextBoxes(i)
                tb.Text = ""
                tb.BackColor = Color.LightPink
            Next
            TotalMilesTextBox.Text = ""
            MileageChargeTextBox.Text = ""
            DayChargeTextBox.Text = ""
            TotalDiscountTextBox.Text = ""
            TotalChargeTextBox.Text = ""
            MilesradioButton.Checked = True
            AAAcheckbox.Checked = False
            Seniorcheckbox.Checked = False

        End If

    End Sub

    'Menu strip clear button
    Private Sub ClearToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem1.Click
        ClearButton_Click(sender, e)
    End Sub

    'Calculate button
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        'What to display in a message to the user
        Dim messageInfo As String = ""
        'Creates a new variable to work with to reduce using the same "inputTextBoxes(i)" over and over again and for validation
        Dim validInput As Boolean
        Dim allValid As Boolean = True
        'Adds all input text boxs to an array to work with
        Dim inputTextBoxes() As TextBox = {NameTextBox, AddressTextBox, CityTextBox, StateTextBox, ZipCodeTextBox,
            BeginOdometerTextBox, EndOdometerTextBox, DaysTextBox}
        Dim tb As TextBox

        'Checks the user has input at least something in all fields from the end to the beginning tab order
        For i As Integer = inputTextBoxes.Count - 1 To 0 Step -1

            'Resets appropriate variables
            tb = inputTextBoxes(i)
            validInput = True

            'Checks to make sure there is a valid value in the user's input for that specific text box, if not, display and hightlight
            If tb.Text.Trim().Length = 0 Then
                validInput = False
            ElseIf tb Is BeginOdometerTextBox Or tb Is EndOdometerTextBox Then
                Try
                    If CInt(BeginOdometerTextBox.Text) > CInt(EndOdometerTextBox.Text) Or CInt(BeginOdometerTextBox.Text) < 0 Then
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
                allValid = False
            End If

        Next

        'Displays a message box if the input data is wrong
        If Not allValid Then
            MessageBox.Show($"Identified Issues:{vbCrLf}{messageInfo}")
        End If

        'If all the inputs are valid and hasn't made an error than it will continue on with the calculations
        If allValid Then

            'Checks if the input values are in miles or kilometers and procceeds with the calculations
            If KilometersradioButton.Checked Then
                TotalMilesTextBox.Text = CStr(0.62 * (CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text))) 'Distance driven in Mi. converted from Km.
            Else
                TotalMilesTextBox.Text = CStr((CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text))) 'Distance driven in Mi.
            End If

            'Applies mileage charge depending on what the reading is.
            If CInt(TotalMilesTextBox.Text) > 500 Then 'miles > 500
                MileageChargeTextBox.Text = Format((CInt(TotalMilesTextBox.Text) * 0.1), "Currency")
            ElseIf CInt(TotalMilesTextBox.Text) >= 201 And CInt(TotalMilesTextBox.Text) <= 500 Then '200 < miles < 500
                MileageChargeTextBox.Text = Format((CInt(TotalMilesTextBox.Text) * 0.12), "Currency")
            Else 'Assumes miles < 200
                MileageChargeTextBox.Text = Format(0, "Currency")
            End If

            'Submits the days and initial total charge
            DayChargeTextBox.Text = Format((CInt(DaysTextBox.Text) * 15), "Currency")
            TotalChargeTextBox.Text = Format(((CDbl(MileageChargeTextBox.Text) + CInt(DayChargeTextBox.Text))), "Currency")

            'Applies Discounts
            If AAAcheckbox.Checked And Not Seniorcheckbox.Checked Then 'If AAA, find total and subtract the discount
                TotalDiscountTextBox.Text = Format((0.05 * CDbl(TotalChargeTextBox.Text)), "Currency")
                TotalChargeTextBox.Text = Format((0.95 * CDbl(TotalChargeTextBox.Text)), "Currency")
            ElseIf Seniorcheckbox.Checked And Not AAAcheckbox.Checked Then ' If Senior, find total and subtract the discount
                TotalDiscountTextBox.Text = Format(((0.03 * CDbl(TotalChargeTextBox.Text))), "Currency")
                TotalChargeTextBox.Text = Format((0.97 * CDbl(TotalChargeTextBox.Text)), "Currency")
            ElseIf Seniorcheckbox.Checked And AAAcheckbox.Checked Then 'If Senior and AAA, find total and subtract the discount
                TotalDiscountTextBox.Text = Format(((0.08 * CDbl(TotalChargeTextBox.Text))), "Currency")
                TotalChargeTextBox.Text = Format((0.92 * CDbl(TotalChargeTextBox.Text)), "Currency")
            ElseIf Not AAAcheckbox.Checked And Not Seniorcheckbox.Checked Then 'If neither of these are checked
                TotalDiscountTextBox.Text = Format(0, "Currency")
            End If 'If there is no discount

            'Adds to the summary
            sumCustomers += 1
            sumMilesDriven = sumMilesDriven + CInt(DaysTextBox.Text)
            sumCharges = sumCharges + CDbl(TotalChargeTextBox.Text)

        End If

    End Sub

    'Menu strip calculate button
    Private Sub CalculateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateToolStripMenuItem.Click
        CalculateButton_Click(sender, e)
    End Sub

    'Summary button
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click

        'Formats and displays totals
        MessageBox.Show($"Total Customers: {sumCustomers}{vbCrLf}Total Miles Driven: {sumMilesDriven}
Total Charges: {Format(sumCharges, "Currency")}", "Detailed Summary?", MessageBoxButtons.OK)

    End Sub

    'Menu strip summary button
    Private Sub SummaryToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SummaryToolStripMenuItem1.Click
        SummaryButton_Click(sender, e)
    End Sub

End Class
