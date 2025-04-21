'ChristopherZ
'Spring 2025
'RCET2265
'Adress Label
'https://github.com/Christopher-isu/CarRental.git


Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm

    ' Summary Variables
    Private totalCustomers As Integer = 0
    Private totalDistance As Double = 0
    Private totalCharges As Double = 0

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Programmatically generate tooltips for all text boxes
        Dim toolTip As New ToolTip()

        toolTip.SetToolTip(NameTextBox, "Enter the customer's name (letters only).")
        toolTip.SetToolTip(AddressTextBox, "Enter the customer's address.")
        toolTip.SetToolTip(CityTextBox, "Enter the city (letters only).")
        toolTip.SetToolTip(StateTextBox, "Enter the state abbreviation (2 letters).")
        toolTip.SetToolTip(ZipCodeTextBox, "Enter the ZIP code (numbers only).")
        toolTip.SetToolTip(BeginOdometerTextBox, "Enter the beginning odometer reading (numeric value).")
        toolTip.SetToolTip(EndOdometerTextBox, "Enter the ending odometer reading (numeric value).")
        toolTip.SetToolTip(DaysTextBox, "Enter the number of days rented (1–45).")
        toolTip.SetToolTip(TotalMilesTextBox, "Automatically calculated: total miles driven.")
        toolTip.SetToolTip(MileageChargeTextBox, "Automatically calculated: mileage charge.")
        toolTip.SetToolTip(DayChargeTextBox, "Automatically calculated: daily charge.")
        toolTip.SetToolTip(TotalDiscountTextBox, "Automatically calculated: total discount.")
        toolTip.SetToolTip(TotalChargeTextBox, "Automatically calculated: total charges.")

        ' Ensure Summary Button starts as disabled
        SummaryButton.Enabled = False
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Try
            ' Input Validation
            Dim errorMessage As String = String.Empty

            If String.IsNullOrWhiteSpace(NameTextBox.Text) OrElse Not IsTextValid(NameTextBox.Text) Then
                errorMessage &= "Name must contain letters only." & vbCrLf
            End If
            If String.IsNullOrWhiteSpace(AddressTextBox.Text) Then
                errorMessage &= "Address is required." & vbCrLf
            End If
            If String.IsNullOrWhiteSpace(CityTextBox.Text) OrElse Not IsTextValid(CityTextBox.Text) Then
                errorMessage &= "City must contain letters only." & vbCrLf
            End If
            If String.IsNullOrWhiteSpace(StateTextBox.Text) OrElse Not IsStateValid(StateTextBox.Text) Then
                errorMessage &= "State must be a 2-letter abbreviation." & vbCrLf
            End If
            If String.IsNullOrWhiteSpace(ZipCodeTextBox.Text) OrElse Not IsNumericValid(ZipCodeTextBox.Text) Then
                errorMessage &= "ZIP code must contain numbers only." & vbCrLf
            End If
            If String.IsNullOrWhiteSpace(BeginOdometerTextBox.Text) OrElse Not IsNumericValid(BeginOdometerTextBox.Text) Then
                errorMessage &= "Beginning odometer reading must be numeric." & vbCrLf
            End If
            If String.IsNullOrWhiteSpace(EndOdometerTextBox.Text) OrElse Not IsNumericValid(EndOdometerTextBox.Text) Then
                errorMessage &= "Ending odometer reading must be numeric." & vbCrLf
            End If
            If String.IsNullOrWhiteSpace(DaysTextBox.Text) OrElse Not IsNumericValid(DaysTextBox.Text) Then
                errorMessage &= "Days rented must be numeric between 1 and 45." & vbCrLf
            End If

            If Not String.IsNullOrEmpty(errorMessage) Then
                MessageBox.Show($"Input Validation Failed:{vbCrLf}{errorMessage}", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Parse numeric values
            Dim beginningOdometer As Double = Double.Parse(BeginOdometerTextBox.Text)
            Dim endingOdometer As Double = Double.Parse(EndOdometerTextBox.Text)
            Dim daysRented As Integer = Integer.Parse(DaysTextBox.Text)

            ' Validate business logic
            If beginningOdometer >= endingOdometer Then
                MessageBox.Show("Beginning odometer reading must be less than ending odometer reading.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            If daysRented <= 0 OrElse daysRented > 45 Then
                MessageBox.Show("Days rented must be greater than 0 and less than or equal to 45.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Calculate distance
            Dim distanceDriven As Double = endingOdometer - beginningOdometer
            If KilometersradioButton.Checked Then
                distanceDriven *= 0.62 ' Convert kilometers to miles
            End If

            ' Perform calculations
            Dim mileageCharge As Double = CalculateMileageCharge(distanceDriven)
            Dim dailyCharge As Double = daysRented * 15.0
            Dim totalBeforeDiscount As Double = dailyCharge + mileageCharge
            Dim discount As Double = CalculateDiscount(AAAcheckbox.Checked, Seniorcheckbox.Checked, totalBeforeDiscount)
            Dim totalCharge As Double = totalBeforeDiscount - discount

            ' Display Outputs
            TotalMilesTextBox.Text = $"{distanceDriven:F2} mi"
            MileageChargeTextBox.Text = $"{mileageCharge:C}"
            DayChargeTextBox.Text = $"{dailyCharge:C}"
            TotalDiscountTextBox.Text = $"{discount:C}"
            TotalChargeTextBox.Text = $"{totalCharge:C}"

            ' Update Summary
            totalCustomers += 1
            totalDistance += distanceDriven
            totalCharges += totalCharge

            ' Enable Summary Button
            SummaryButton.Enabled = True

        Catch ex As Exception
            MessageBox.Show($"An error occurred while processing your input: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        ' Clear Input Fields
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()

        ' Clear Output Fields
        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()

        ' Reset Checkboxes and Radio Buttons
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        ' Display Summary
        MessageBox.Show($"Summary:{vbCrLf}" &
                        $"Total Customers: {totalCustomers}{vbCrLf}" &
                        $"Total Distance Driven: {totalDistance:F2} mi{vbCrLf}" &
                        $"Total Charges: {totalCharges:C}",
                        "Summary", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        ' Exit Confirmation
        Dim response = MessageBox.Show("Are you sure you want to exit?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If response = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    ' Validation Helper Methods
    Private Function IsTextValid(input As String) As Boolean
        Return Not String.IsNullOrWhiteSpace(input) AndAlso input.All(AddressOf Char.IsLetter)
    End Function

    Private Function IsStateValid(input As String) As Boolean
        Return input.Length = 2 AndAlso input.All(AddressOf Char.IsLetter)
    End Function

    Private Function IsNumericValid(input As String) As Boolean
        Dim number As Double
        Return Double.TryParse(input, number)
    End Function

    Private Function CalculateMileageCharge(distance As Double) As Double
        If distance <= 200 Then
            Return 0
        ElseIf distance <= 500 Then
            Return (distance - 200) * 0.12
        Else
            Return (300 * 0.12) + ((distance - 500) * 0.1)
        End If
    End Function

    Private Function CalculateDiscount(isAAA As Boolean, isSenior As Boolean, total As Double) As Double
        Dim discount As Double = 0
        If isAAA Then
            discount += total * 0.05
        End If
        If isSenior Then
            discount += total * 0.03
        End If
        Return discount
    End Function

End Class
