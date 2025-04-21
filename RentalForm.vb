Option Explicit On
Option Strict On
Option Compare Binary

Imports System

Public Class RentalForm

    ' Summary Variables
    Private totalCustomers As Integer = 0
    Private totalDistance As Double = 0
    Private totalCharges As Double = 0

    Public Sub StartRentalForm()
        Dim exitProgram As Boolean = False

        While Not exitProgram
            Console.Clear()
            Console.WriteLine("Car Rental Application")

            ' Input Section
            Console.Write("Enter Customer Name: ")
            Dim customerName As String = Console.ReadLine()

            Console.Write("Beginning Odometer Reading: ")
            Dim beginningOdometer As Double = ValidateDoubleInput()

            Console.Write("Ending Odometer Reading: ")
            Dim endingOdometer As Double = ValidateDoubleInput()

            While beginningOdometer >= endingOdometer
                Console.WriteLine("Error: Beginning odometer must be less than ending odometer.")
                Console.Write("Re-enter Ending Odometer Reading: ")
                endingOdometer = ValidateDoubleInput()
            End While

            Console.Write("Number of Days Rented: ")
            Dim daysRented As Integer = ValidateIntegerInput()

            While daysRented <= 0 OrElse daysRented > 45
                Console.WriteLine("Error: Days must be greater than 0 and no more than 45.")
                Console.Write("Re-enter Number of Days Rented: ")
                daysRented = ValidateIntegerInput()
            End While

            Console.WriteLine("Enter 1 for Miles or 2 for Kilometers:")
            Dim unit As Integer = ValidateIntegerInput()

            While unit <> 1 AndAlso unit <> 2
                Console.WriteLine("Error: Enter 1 for Miles or 2 for Kilometers.")
                unit = ValidateIntegerInput()
            End While

            ' Discount Section
            Console.Write("AAA Member (yes/no): ")
            Dim aaaMember As Boolean = (Console.ReadLine().ToLower() = "yes")

            Console.Write("Senior Citizen (yes/no): ")
            Dim seniorCitizen As Boolean = (Console.ReadLine().ToLower() = "yes")

            ' Conversion and Calculation
            Dim distanceDriven As Double = endingOdometer - beginningOdometer
            If unit = 2 Then
                distanceDriven *= 0.62 ' Convert kilometers to miles
            End If

            Dim mileageCharge As Double = CalculateMileageCharge(distanceDriven)
            Dim dailyCharge As Double = daysRented * 15.0
            Dim discount As Double = CalculateDiscount(aaaMember, seniorCitizen, dailyCharge + mileageCharge)
            Dim totalCharge As Double = (dailyCharge + mileageCharge) - discount

            ' Output Section
            Console.WriteLine($"Distance Driven: {distanceDriven:F2} mi")
            Console.WriteLine($"Mileage Charge: {mileageCharge:C}")
            Console.WriteLine($"Daily Charge: {dailyCharge:C}")
            Console.WriteLine($"Discount: {discount:C}")
            Console.WriteLine($"Total Charge: {totalCharge:C}")

            ' Update Summary
            totalCustomers += 1
            totalDistance += distanceDriven
            totalCharges += totalCharge

            ' Summary or Exit Option
            Console.Write("Type 'summary' for summary or 'exit' to quit: ")
            Dim userChoice As String = Console.ReadLine().ToLower()
            If userChoice = "summary" Then
                DisplaySummary()
            ElseIf userChoice = "exit" Then
                exitProgram = ConfirmExit()
            End If
        End While
    End Sub

    Private Function ValidateDoubleInput() As Double
        Dim input As String
        Dim value As Double

        Do
            input = Console.ReadLine()
            If Double.TryParse(input, value) Then
                Return value
            End If
            Console.Write("Invalid input. Please enter a valid number: ")
        Loop
    End Function

    Private Function ValidateIntegerInput() As Integer
        Dim input As String
        Dim value As Integer

        Do
            input = Console.ReadLine()
            If Integer.TryParse(input, value) Then
                Return value
            End If
            Console.Write("Invalid input. Please enter a valid integer: ")
        Loop
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

    Private Function CalculateDiscount(aaa As Boolean, senior As Boolean, total As Double) As Double
        Dim discount As Double = 0

        If aaa Then
            discount += total * 0.05
        End If

        If senior Then
            discount += total * 0.03
        End If

        Return discount
    End Function

    Private Sub DisplaySummary()
        Console.WriteLine("Summary")
        Console.WriteLine($"Total Customers: {totalCustomers}")
        Console.WriteLine($"Total Distance Driven: {totalDistance:F2} mi")
        Console.WriteLine($"Total Charges: {totalCharges:C}")
    End Sub

    Private Function ConfirmExit() As Boolean
        Console.Write("Are you sure you want to exit? (yes/no): ")
        Return Console.ReadLine().ToLower() = "yes"
    End Function

End Class

' Create an instance and start the application
Module Program
    Sub Main()
        Dim rentalForm As New RentalForm()
        rentalForm.StartRentalForm()
    End Sub
End Module
