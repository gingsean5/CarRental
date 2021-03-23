Option Explicit On
Option Strict On
Option Compare Binary
Imports System.ComponentModel

Public Class RentalForm

    Function ValidateFields() As String
        Dim RMessage As String
        Dim AllMessage As String
        Dim DayInt As Integer
        If DaysTextBox.Text = "" Then
            AllMessage = "All fields are required"
            DaysTextBox.Focus()
        Else
            Try
                DayInt = CInt(DaysTextBox.Text)
            Catch ex As Exception
                RMessage &= "Number of Days must be a number" & vbNewLine
                DaysTextBox.Text = ""
                DaysTextBox.Focus()
            End Try
            If DayInt <= 0 Then
                RMessage &= "Number of Days must be more than 0" & vbNewLine
                DaysTextBox.Text = ""
                DaysTextBox.Focus()
            ElseIf DayInt > 45 Then
                RMessage &= "Number of Days cannot be more than 45" & vbNewLine
                DaysTextBox.Text = ""
                DaysTextBox.Focus()
            End If
        End If

        Dim EndOInt As Integer
        If EndOdometerTextBox.Text = "" Then
            AllMessage = "All fields are required"
            EndOdometerTextBox.Focus()
        Else
            Try
                EndOInt = CInt(EndOdometerTextBox.Text)
            Catch ex As Exception
                RMessage &= "Ending Odometer Reading must be a number" & vbNewLine
                EndOdometerTextBox.Text = ""
                EndOdometerTextBox.Focus()
            End Try
        End If

        Dim BeginOInt As Integer
        If BeginOdometerTextBox.Text = "" Then
            AllMessage = "All fields are required"
            BeginOdometerTextBox.Focus()
        Else
            Try
                BeginOInt = CInt(BeginOdometerTextBox.Text)
            Catch ex As Exception
                RMessage &= "Beginning Odometer Reading must be a number" & vbNewLine
                BeginOdometerTextBox.Text = ""
                BeginOdometerTextBox.Focus()
            End Try
        End If
        If BeginOInt > EndOInt Then
            RMessage &= "Beginning odometer reading cannot be larger than the ending odometer reading" & vbNewLine
            BeginOdometerTextBox.Text = ""
            EndOdometerTextBox.Text = ""
            BeginOdometerTextBox.Focus()
        End If

        Dim ZipInt As Integer
        If ZipCodeTextBox.Text = "" Then
            AllMessage = "All fields are required"
            ZipCodeTextBox.Focus()
        Else
            Try
                ZipInt = CInt(ZipCodeTextBox.Text)
            Catch ex As Exception
                RMessage &= "Zipcode must be a number"
                ZipCodeTextBox.Text = ""
                ZipCodeTextBox.Focus()
            End Try
        End If

        If StateTextBox.Text = "" Then
            AllMessage = "All Fields are required"
            StateTextBox.Focus()
        End If

        If CityTextBox.Text = "" Then
            AllMessage = "All Fields are required"
            CityTextBox.Focus()
        End If

        If AddressTextBox.Text = "" Then
            AllMessage = "All Fields are required"
            AddressTextBox.Focus()
        End If

        If NameTextBox.Text = "" Then
            AllMessage = "All Fields are required"
            NameTextBox.Focus()
        End If

        Dim returnMessage As String
        If AllMessage = "" And RMessage = "" Then
            returnMessage = ""
        Else
            returnMessage = AllMessage & vbNewLine & RMessage
        End If


        Return returnMessage
    End Function

    Function MilesCharge() As Double


        Dim MileCount As Integer
        Dim MileCharge As Double
        Dim MileCharge0 As Double
        Dim MileCharge200 As Double
        Dim MileCharge500 As Double
        Dim Cost As Double

        If ValidateFields() = "" Then
            Dim EndOInt As Integer
            EndOInt = CInt(EndOdometerTextBox.Text)
            Dim BeginOInt As Integer
            BeginOInt = CInt(BeginOdometerTextBox.Text)
            If MilesradioButton.Checked = True Then
                MileCount = EndOInt - BeginOInt
            End If
            If KilometersradioButton.Checked = True Then
                MileCount = CInt((EndOInt - BeginOInt) * 0.62)
            End If
            If MileCount <= 200 Then
                MileCharge0 = 0
            ElseIf MileCount > 200 And MileCount <= 500 Then
                MileCharge200 = 0.12 * (MileCount - 200)
            ElseIf MileCount > 500 Then
                MileCharge200 = 0.12 * 300
                MileCharge500 = 0.1 * (MileCount - 500)
            End If
            MileCharge = MileCharge0 + MileCharge200 + MileCharge500

            Cost = MileCharge
        End If
        Return Cost
    End Function

    Function MileCount() As Integer
        Dim count As Integer
        If ValidateFields() = "" Then
            Dim EndOInt As Integer
            EndOInt = CInt(EndOdometerTextBox.Text)
            Dim BeginOInt As Integer
            BeginOInt = CInt(BeginOdometerTextBox.Text)
            count = EndOInt - BeginOInt
        End If
        Return count
    End Function

    Function DayCharge() As Double
        Dim DailyCharge As Double
        Dim DayInt As Integer
        If ValidateFields() = "" Then
            DayInt = CInt(DaysTextBox.Text)
            DailyCharge = DayInt * 15
            DayChargeTextBox.Text = CStr(DailyCharge)

        End If

        Return DailyCharge
    End Function

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click, CalculateToolStripMenuItem.Click
        Dim ProblemMessages As String
        Dim MileCharge As Double
        Dim MileString As String
        Dim DailyCharge As Double
        Dim DayString As String
        Dim TotalCost As Double
        Dim MilesDriven As Integer
        Dim KilosDriven As Integer
        Dim CostPreDiscount As Double
        Dim AAADiscount As Double
        Dim AAAstring As String
        Dim SeniorDiscount As Double
        Dim SeniorString As String
        Dim TotalDiscount As Double
        Dim TotalStringDiscount As String
        Dim FinalCost As Double
        Dim FinalString As String

        ProblemMessages = ValidateFields()
        If ProblemMessages <> "" Then
            MsgBox(ProblemMessages)
        Else
            SummaryButton.Enabled = True
            DailyCharge = DayCharge()
            DayString = FormatCurrency(DailyCharge, , , TriState.True, TriState.True)
            DayChargeTextBox.Text = DayString
            MileCharge = MilesCharge()
            MileString = FormatCurrency(MileCharge, , , TriState.True, TriState.True)
            MileageChargeTextBox.Text = MileString

            TotalCost = MileCharge + DailyCharge

            MilesDriven = MileCount()
            KilosDriven = CInt(0.62 * MilesDriven)
            If MilesradioButton.Checked = True Then
                TotalMilesTextBox.Text = CStr(MilesDriven) + " mi"
            End If
            If KilometersradioButton.Checked = True Then
                TotalMilesTextBox.Text = CStr(KilosDriven) + " mi"

            End If
            CostPreDiscount = DailyCharge + MileCharge
            If AAAcheckbox.Checked = True Or Seniorcheckbox.Checked = True Then
                If AAAcheckbox.Checked = True Then
                    AAADiscount = CostPreDiscount * 0.05
                    AAAstring = FormatCurrency(AAADiscount, , , TriState.True, TriState.True)

                End If
                If Seniorcheckbox.Checked = True Then
                    SeniorDiscount = CostPreDiscount * 0.03
                    SeniorString = FormatCurrency(SeniorDiscount, , , TriState.True, TriState.True)
                End If
                TotalDiscount = AAADiscount + SeniorDiscount
                TotalStringDiscount = FormatCurrency(TotalDiscount, , , TriState.True, TriState.True)
                TotalDiscountTextBox.Text = TotalStringDiscount
            Else
                TotalDiscountTextBox.Text = "$0.00"
            End If

            FinalCost = TotalCost - TotalDiscount
            FinalString = FormatCurrency(FinalCost, , , TriState.True, TriState.True)
            TotalChargeTextBox.Text = FinalString
        End If
    End Sub

    Function CleartheForm() As Boolean
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        DaysTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""

        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""

        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Function

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem1.Click
        CleartheForm()
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem1.Click

        Me.Close()
    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        SummaryButton.Enabled = False

    End Sub

    Function Summary() As String
        Dim SummaryMessage As String


        SummaryMessage = $"Total Customers:     
Total Miles Driven:     
Total Charges:      "
        Return SummaryMessage
    End Function

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click, SummaryToolStripMenuItem1.Click

        MsgBox(Summary())

        CleartheForm()
    End Sub

    Private Sub RentalForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dim words As String = "Are you sure that you would like to close the form?"
        Dim caption As String = "Form Closing"
        Dim result As DialogResult
        result = MessageBox.Show(words, caption,
                                 MessageBoxButtons.YesNo,
                                 MessageBoxIcon.Question)
        If result = DialogResult.No Then
            e.Cancel = True
        End If
    End Sub


End Class
