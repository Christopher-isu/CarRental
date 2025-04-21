'ChristopherZ
'Spring 2025
'RCET2265
'Adress Label
'https://github.com/Christopher-isu/CarRental.git


Option Explicit On
Option Strict On
Option Compare Binary
'test
Public Class RentalForm

    ' Summary Variables
    Private totalCustomers As Integer = 0
    Private totalDistance As Double = 0
    Private totalCharges As Double = 0

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        ' Input Validation
        If String.IsNullOrWhiteSpace(NameTextBox.Text) OrElse
           String.IsNullOrWhiteSpace(BeginOdometerTextBox.Text) OrElse
           String.IsNullOrWhiteSpace(EndOdometerTextBox.Text) OrElse
           String.IsNullOrWhiteSpace(DaysTextBox.Text) Then

            MessageBox.Show("All fields must be filled out.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim beginningOdometer, endingOdometer, distanceDriven As Double
        Dim daysRented As Integer

        If Not Double.TryParse(BeginOdometerTextBox.Text, beginningOdometer) OrElse
           Not Double.TryParse(EndOdometerTextBox.Text, endingOdometer) OrElse
           Not Integer.TryParse(DaysTextBox.Text, daysRented) Then

            MessageBox.Show("Please enter valid numeric values.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If beginningOdometer >= endingOdometer Then
            MessageBox.Show("Beginning odometer reading must be less than ending odometer reading.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If daysRented <= 0 OrElse daysRented > 45 Then
            MessageBox.Show("Days rented must be greater than 0 and less than or equal to 45.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Calculate Distance
        distanceDriven = endingOdometer - beginningOdometer
        If KilometersradioButton.Checked Then
            distanceDriven *= 0.62 ' Convert kilometers to miles
        End If

        ' Calculations
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
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        ' Clear Input Fields
        NameTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()

        ' Clear Address Fields
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()

        ' Clear Output Labels
        TotalMilesTextBox.Text = String.Empty
        MileageChargeTextBox.Text = String.Empty
        DayChargeTextBox.Text = String.Empty
        TotalDiscountTextBox.Text = String.Empty
        TotalChargeTextBox.Text = String.Empty

        ' Reset Checkboxes and Radio Buttons
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        MilesradioButton.Checked = True
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        ' Display Summary
        MessageBox.Show($"Summary:" & vbCrLf &
                        $"Total Customers: {totalCustomers}" & vbCrLf &
                        $"Total Distance Driven: {totalDistance:F2} mi" & vbCrLf &
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

    ' Helper Methods
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

    Private Sub TotalDiscountTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalDiscountTextBox.TextChanged

    End Sub
End Class