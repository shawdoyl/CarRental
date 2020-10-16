'Shaw_Doyle
'RCET0265
'Fall 2020
'Car_Rental_Form
'https://github.com/shawdoyl/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Dim numCustomers As Integer
    Dim totalMiles, totalCharges, totalDiscount, mileCharge As Decimal

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ResetAll()
        SummaryButton.Enabled = False
        'Preset inputs for testing
        'NameTextBox.Text = "Doyle Shaw"
        'AddressTextBox.Text = "2720 E 95th N"
        'CityTextBox.Text = "Idaho Falls"
        'StateTextBox.Text = "Idaho"
        'ZipCodeTextBox.Text = "83401"
        'BeginOdometerTextBox.Text = "123456"
        'EndOdometerTextBox.Text = "123689"
        'DaysTextBox.Text = "15"
    End Sub
    Sub ResetAll()
        'Clear User input and output
        'Inputs
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        MilesradioButton.Select()
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        'Outputs
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
    End Sub
    Function MileageCharge(ByRef miles As Decimal) As Decimal
        'First 200 miles driven are always free. 
        'All miles between 201 And 500 are .12 cents per mile. 
        'miles greater than 500 are charged at .10 cents per mile.
        Const RATEREGULAR = 0.12D
        Const RATELOW = 0.1D
        Const RATEFREE = 0D

        Select Case miles
            Case <= 200
                mileCharge = miles * RATEFREE
            Case > 500
                mileCharge = (300) * RATEREGULAR
                mileCharge += (miles - 500) * RATELOW
            Case Else
                mileCharge = (miles - 200) * RATEREGULAR
        End Select
        Return mileCharge
    End Function
    Function Discount(totalCharges As Decimal) As Decimal
        'AAA members receive a 5% discount
        'senior citizens get a 3% discount.
        Const AAARATE = 0.05D
        Const SENIORRATE = 0.03D

        If AAAcheckbox.Checked = True Then
            totalDiscount += totalCharges * AAARATE
        End If
        If Seniorcheckbox.Checked = True Then
            totalDiscount += totalCharges * SENIORRATE
        End If
        Return totalDiscount
    End Function
    Function UserMessages(addMessage As Boolean, message As String, clearMessage As Boolean) As String
        Static formattedMessages As String

        If clearMessage = True Then
            formattedMessages = ""
        ElseIf addMessage = True Then
            formattedMessages &= message & vbNewLine
        End If
        Return formattedMessages
    End Function
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim dataValidated As Boolean = False
        Dim beginOdometerGood As Boolean = False
        Dim endOdometerGood As Boolean = False
        Dim daysNumberGood As Boolean = False
        Dim beginOdometer, endOdometer As Decimal
        Dim numberOfDays As Integer
        Dim userMessage As String = ""

        If NameTextBox.Text = "" Then
            userMessage = "Please enter a Name." & vbNewLine
        End If
        If AddressTextBox.Text = "" Then
            userMessage &= "Please enter the Address." & vbNewLine
        End If
        If CityTextBox.Text = "" Then
            userMessage &= "Please enter the City." & vbNewLine
        End If
        If StateTextBox.Text = "" Then
            userMessage &= "Please enter the State." & vbNewLine
        End If
        If ZipCodeTextBox.Text = "" Then
            userMessage &= "Please enter the Zip Code." & vbNewLine
        End If
        If BeginOdometerTextBox.Text = "" Then
            userMessage &= "Please enter the beginning odometer reading." & vbNewLine
        End If
        If EndOdometerTextBox.Text = "" Then
            userMessage &= "Please enter the end odometer reading." & vbNewLine
        End If
        If DaysTextBox.Text = "" Then
            userMessage &= "Please enter the number of days." & vbNewLine
        End If
        'if any of the last 3 are not blank, check to see if they are valid integers
        If BeginOdometerTextBox.Text <> "" Or EndOdometerTextBox.Text <> "" Or
            DaysTextBox.Text <> "" Then
            Try
                beginOdometer = CDec(BeginOdometerTextBox.Text)
                beginOdometerGood = True
            Catch ex As Exception
                'did not convert to dec
                userMessage &= "Please enter a valid beginning odometer reading" & vbNewLine
            End Try
            Try
                endOdometer = CDec(EndOdometerTextBox.Text)
                endOdometerGood = True
            Catch ex As Exception
                userMessage &= "Please enter a valid ending odometer reading" & vbNewLine
            End Try
            If beginOdometerGood = True And endOdometerGood = True Then
                'as long as both values converted to dec, compare them
                If beginOdometer < endOdometer Then
                    'do nothing. Beginning is less than ending
                Else
                    userMessage &= "The ending odometer reading must be greater than the 
                    starting odometer reading. Please fix this to continue." & vbNewLine _
                    & vbNewLine
                End If
            End If
            Try
                numberOfDays = CInt(DaysTextBox.Text)
                If numberOfDays > 0 And numberOfDays <= 45 Then
                    'days is ok
                    daysNumberGood = True
                Else
                    'if this line is entered, days input is a number but not in range
                    userMessage &= "The number of days must be between 0 and 45." & vbNewLine
                End If
            Catch ex As Exception
                userMessage &= "Please enter a valid number for the amount of days." _
                    & vbNewLine
            End Try
        End If
        If userMessage <> "" Then
            MsgBox(userMessage)
            'following 3 ifs clears bad data
            If beginOdometerGood = False Then
                BeginOdometerTextBox.Text = ""
            End If
            If endOdometerGood = False Then
                EndOdometerTextBox.Text = ""
            End If
            If daysNumberGood = False Then
                DaysTextBox.Text = ""
            End If
            'Following ifs determines which bad data to select (last instance is selected)
            'if statements for each text box, from name to days. 
            'If bad, Then NameTextBox.Select()
            If daysNumberGood = False Then
                DaysTextBox.Select()
            End If
            If endOdometerGood = False Then
                'endOdometerGood corresponds to validation of end odometer text
                EndOdometerTextBox.Select()
            End If
            If beginOdometerGood = False Then
                'beginOdometerGood corresponds to validation of begin odometer text
                BeginOdometerTextBox.Select()
            End If
            If ZipCodeTextBox.Text = "" Then
                ZipCodeTextBox.Select()
            End If
            If StateTextBox.Text = "" Then
                StateTextBox.Select()
            End If
            If CityTextBox.Text = "" Then
                CityTextBox.Select()
            End If
            If AddressTextBox.Text = "" Then
                AddressTextBox.Select()
            End If
            If NameTextBox.Text = "" Then
                NameTextBox.Select()
            End If
        Else
            dataValidated = True
        End If
        If dataValidated = True Then
            'validation checks out. do math here
            Dim milesDriven, milesDrivenRounded, totalOwed, subTotal As Decimal
            Dim dayCharge As Integer

            If MilesradioButton.Checked = True Then
                milesDriven = endOdometer - beginOdometer
            Else
                milesDriven = (endOdometer - beginOdometer) * 0.62D
            End If
            milesDrivenRounded = Math.Round(milesDriven, 2)
            MileageCharge(milesDrivenRounded)
            'daily charge =  days * $15
            dayCharge = numberOfDays * 15
            DayChargeTextBox.Text = dayCharge.ToString("C")
            subTotal = mileCharge + dayCharge
            TotalMilesTextBox.Text = CStr(milesDrivenRounded)
            MileageChargeTextBox.Text = CStr(mileCharge)
            Discount(subTotal)
            TotalDiscountTextBox.Text = totalDiscount.ToString("c")
            totalOwed = mileCharge + dayCharge - totalDiscount
            TotalChargeTextBox.Text = totalOwed.ToString("c")
            'Add to the number of customers, and total charges and miles
            numCustomers += 1
            totalMiles += milesDriven
            totalCharges += totalOwed
            SummaryButton.Enabled = True
        End If
    End Sub
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Dim totalSummary As String = ""
        totalSummary = "Total Customers:       " & numCustomers.ToString & vbNewLine
        totalSummary &= "Total Miles Driven:    " & totalMiles.ToString("n") & " mi" & vbNewLine
        totalSummary &= "Total charges:           " & totalCharges.ToString("c") & vbNewLine
        MessageBox.Show(totalSummary, "Detailed Summary")
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        ResetAll()
        MileageChargeTextBox.Clear()
        TotalMilesTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()
    End Sub
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim Msg, Title As String
        Dim Style As MsgBoxStyle
        Dim Response As Integer

        Msg = "Do you want to close the program?"
        Title = "Close Program?"
        Style = CType(vbYesNo + vbCritical + vbDefaultButton2, MsgBoxStyle)
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            Msg = "Are you sure you want to close the program?"
            Title = "Are You Sure?"
            Style = CType(vbYesNo + vbCritical + vbDefaultButton2, MsgBoxStyle)
            Response = MsgBox(Msg, Style, Title)
            If Response = vbYes Then
                Msg = "Are you really sure you want to close the program?"
                Title = "Are You Really Sure?"
                Style = CType(vbYesNo + vbCritical + vbDefaultButton2, MsgBoxStyle)
                Response = MsgBox(Msg, Style, Title)
                If Response = vbYes Then
                    Me.Close()
                Else
                End If
            Else
            End If
        Else
        End If
    End Sub
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Car Rental program Beta Version 1.0.001" & vbNewLine _
               & "Doyle Shaw" & vbNewLine _
               & "Fall 2020" & vbNewLine _
               & "RCET0265" & vbNewLine _
               & "In association with Doofenshmirtz Evil Incorporated.")
    End Sub
End Class