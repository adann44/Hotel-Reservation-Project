' Name Project:     Hotel Reservation
' Purpose:          Hotel Reservation Prices for guests
' Names and Date:   Adan Navejas-Gallegos 9/17/22


Option Strict On
Option Explicit On
Option Infer Off

Public Class frmMain

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtKingNights.Enabled = False
        txtKingRooms.Enabled = False

        txtQueenNights.Enabled = False
        txtQueenRooms.Enabled = False

        txtDoubleNights.Enabled = False
        txtDoubleRooms.Enabled = False

        lstSpecialRates.Items.Add("None")
        lstSpecialRates.Items.Add("Corporate")
        lstSpecialRates.Items.Add("AAA/CAA")
        lstSpecialRates.Items.Add("Gov + Military")

        lstSpecialRates.SelectedItem = "None"

        For dblNumChildren As Double = 0 To 6
            lstNumChildren.Items.Add(dblNumChildren)

        Next dblNumChildren
        lstNumChildren.SelectedIndex = 0


        For dblNumAdults As Double = 1 To 20
            lstNumAdults.Items.Add(dblNumAdults)

        Next dblNumAdults
        lstNumAdults.SelectedIndex = 0



    End Sub

    Private Sub chkSelectQueen_CheckedChanged(sender As Object, e As EventArgs) Handles chkSelectQueen.CheckedChanged
        If chkSelectQueen.Checked Then
            txtQueenNights.Enabled = True
            txtQueenRooms.Enabled = True
        Else
            txtQueenNights.Enabled = False
            txtQueenRooms.Enabled = False
            txtQueenNights.Text = String.Empty
            txtQueenRooms.Text = String.Empty
        End If

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty


    End Sub

    Private Sub chkSelectDouble_CheckedChanged(sender As Object, e As EventArgs) Handles chkSelectDouble.CheckedChanged
        If chkSelectDouble.Checked Then
            txtDoubleNights.Enabled = True
            txtDoubleRooms.Enabled = True
        Else
            txtDoubleNights.Enabled = False
            txtDoubleRooms.Enabled = False
            txtDoubleNights.Text = String.Empty
            txtDoubleRooms.Text = String.Empty
        End If

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        ' Declare all Constants. 
        Const dblKingRegular As Double = 160
        Const dblKingMember As Double = 150
        Const dblQueenMember As Double = 170
        Const dblQueenRegular As Double = 180
        Const dblDoubleRegular As Double = 145
        Const dblDoubleMember As Double = 140
        Const dblCorporate As Double = 0.08
        Const dblAAA As Double = 0.06
        Const dblGov As Double = 0.1
        Const dblPromo As Double = 0.15
        Const dblTaxRate As Double = 0.1425
        Const dblRoomFee As Double = 12.5



        ' Declare all variables. 
        Dim dblKingRooms As Double
        Dim dblKingNights As Double
        Dim dblQueenRooms As Double
        Dim dblQueenNights As Double
        Dim dblDoubleRooms As Double
        Dim dblDoubleNights As Double
        Dim dblTotalRoomsBooked As Double
        Dim dblTotalGuestsBooked As Double
        Dim dblTotalNightsBooked As Double
        Dim dblKingRoomCharge As Double
        Dim dblQueenRoomCharge As Double
        Dim dblDoubleRoomCharge As Double
        Dim dblTotalRoomCharge As Double
        Dim dblTax As Double
        Dim dblHotelFees As Double
        Dim dblTotalDue As Double

        ' Convert to numeric data type. 
        Double.TryParse(txtKingRooms.Text, dblKingRooms)
        Double.TryParse(txtKingNights.Text, dblKingNights)
        Double.TryParse(txtQueenRooms.Text, dblQueenRooms)
        Double.TryParse(txtQueenNights.Text, dblQueenNights)
        Double.TryParse(txtDoubleRooms.Text, dblDoubleRooms)
        Double.TryParse(txtDoubleNights.Text, dblDoubleNights)



        'Do While intKingRooms <= 4 And intQueenRooms <= 4 And intDoubleRooms <= 4 
        ' Your code. 

        ' Max FIVE guests per room Message prompt for King alone
        If dblKingRooms = 1 AndAlso lstNumAdults.SelectedIndex > 4 Then
            MsgBox("You have exceeded the maximum guests per room!")

        ElseIf dblKingRooms = 2 AndAlso lstNumAdults.SelectedIndex > 9 Then
            MsgBox("You have exceeded the maximum guests per room!")
        ElseIf dblKingRooms = 3 AndAlso lstNumAdults.SelectedIndex > 14 Then
            MsgBox("You have exceeded the maximum guests per room!")

            'Max FIVE guests per room Message prompt for Queeen alone 
        ElseIf dblQueenRooms = 1 And lstNumAdults.SelectedIndex > 4 Then
            MsgBox("You have exceeded the maximum guests per room!")
        ElseIf dblQueenRooms = 2 And lstNumAdults.SelectedIndex > 9 Then
            MsgBox("You have exceeded the maximum guests per room!")
        ElseIf dblQueenRooms = 3 And lstNumAdults.SelectedIndex > 14 Then
            MsgBox("You have exceeded the maximum guests per room!")

            'Max FIVE guests per room Message prompt for Double alone
        ElseIf dblDoubleRooms = 1 And lstNumAdults.SelectedIndex > 4 Then
            MsgBox("You have exceeded the maximum guests per room!")
        ElseIf dblDoubleRooms = 2 And lstNumAdults.SelectedIndex > 9 Then
            MsgBox("You have exceeded the maximum guests per room!")
        ElseIf dblDoubleRooms = 3 And lstNumAdults.SelectedIndex > 14 Then
            MsgBox("You have exceeded the maximum guests per room!")

        End If


        ' For total rooms booked and total nights. 
        If dblKingRooms > 0 Or dblQueenRooms > 0 Or dblDoubleRooms > 0 Then
            dblTotalRoomsBooked = dblKingRooms + dblQueenRooms + dblDoubleRooms
            If dblKingNights > 0 Or dblQueenNights > 0 Or dblDoubleNights > 0 Then
                dblTotalNightsBooked = dblKingNights + dblQueenNights + dblDoubleNights
            End If
        End If

        'Display total guests 
        ' Your code. 
        Dim TotalGuestsPlaceHolder As Integer = 1
        Dim TotalGuests As Integer
        TotalGuests = TotalGuestsPlaceHolder + lstNumAdults.SelectedIndex + lstNumChildren.SelectedIndex
        'lblTotalGuestsBooked.Text = TotalGuests.ToString

        ' For room charge calculations. 
        If radMemberKing.Checked = True Then
            dblKingRoomCharge = (dblKingRooms) * (dblKingNights) * (dblKingMember)
            dblTotalRoomCharge += dblKingRoomCharge

        End If

        If radRegularKing.Checked = True Then
            dblKingRoomCharge = (dblKingRooms) * (dblKingNights) * (dblKingRegular)
            dblTotalRoomCharge += dblKingRoomCharge

        End If

        If radMemeberQueen.Checked = True Then
            dblQueenRoomCharge = (dblQueenRooms) * (dblQueenNights) * (dblQueenMember)
            dblTotalRoomCharge += dblQueenRoomCharge

        End If

        If radRegularQueen.Checked = True Then
            dblQueenRoomCharge = (dblQueenRooms) * (dblQueenNights) * (dblQueenRegular)
            dblTotalRoomCharge += dblQueenRoomCharge

        End If

        If radMemberDouble.Checked = True Then
            dblDoubleRoomCharge = (dblDoubleRooms) * (dblDoubleNights) * (dblDoubleMember)
            dblTotalRoomCharge += dblDoubleRoomCharge

        End If

        If radRegularDouble.Checked = True Then
            dblDoubleRoomCharge = (dblDoubleRooms) * (dblDoubleNights) * (dblDoubleRegular)
            dblTotalRoomCharge += dblDoubleRoomCharge

        End If

        ' For special rates applied to room charge. 
        Select Case lstSpecialRates.SelectedIndex

            Case 0
                dblTotalRoomCharge = dblTotalRoomCharge

            Case 1
                dblTotalRoomCharge -= (dblTotalRoomCharge * dblCorporate)

            Case 2
                dblTotalRoomCharge -= (dblTotalRoomCharge * dblAAA)

            Case 3
                dblTotalRoomCharge -= (dblTotalRoomCharge * dblGov)

        End Select

        ' To apply the promo code. 
        If txtPromo.Text.Trim.ToUpper = "MISRULES" Then
            dblTotalRoomCharge -= (dblTotalRoomCharge * dblPromo)
            dblTotalDue += dblTotalRoomCharge
        Else dblTotalRoomCharge = dblTotalRoomCharge

        End If

        ' For Tax Calculation. 
        If dblTotalRoomCharge > 0 Then
            dblTax = dblTotalRoomCharge * dblTaxRate
            dblTotalDue += dblTax
        End If

        ' For hotel fees calculation. 
        If dblTotalRoomsBooked > 0 And dblTotalNightsBooked > 0 Then
            dblHotelFees = ((dblKingRooms * dblKingNights) * dblRoomFee) + ((dblQueenRooms * dblQueenNights) * dblRoomFee) + ((dblDoubleRooms * dblDoubleNights) * dblRoomFee)
            dblTotalDue += dblHotelFees
        End If

        ' Converting everything back to string data type. 
        lblTotalRoomsBooked.Text = dblTotalRoomsBooked.ToString
        lblTotalGuestsBooked.Text = dblTotalGuestsBooked.ToString
        lblTotalNightsBooked.Text = dblTotalNightsBooked.ToString
        lblRoomCharge.Text = dblTotalRoomCharge.ToString("C2")
        lblTax.Text = dblTax.ToString("C2")
        lblHotelFees.Text = dblHotelFees.ToString("C2")
        lblTotalDue.Text = dblTotalDue.ToString("C2")

    End Sub

    Private Sub chkSelectKing_CheckedChanged(sender As Object, e As EventArgs) Handles chkSelectKing.CheckedChanged
        If chkSelectKing.Checked Then
            txtKingNights.Enabled = True
            txtKingRooms.Enabled = True
        Else
            txtKingNights.Enabled = False
            txtKingRooms.Enabled = False
            txtKingNights.Text = String.Empty
            txtKingRooms.Text = String.Empty
        End If

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Me.Close()
    End Sub

    Private Sub txtKingRooms_TextChanged(sender As Object, e As EventArgs) Handles txtKingRooms.TextChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub txtKingNights_TextChanged(sender As Object, e As EventArgs) Handles txtKingNights.TextChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub txtQueenRooms_TextChanged(sender As Object, e As EventArgs) Handles txtQueenRooms.TextChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub txtQueenNights_TextChanged(sender As Object, e As EventArgs) Handles txtQueenNights.TextChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub txtDoubleRooms_TextChanged(sender As Object, e As EventArgs) Handles txtDoubleRooms.TextChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub txtDoubleNights_TextChanged(sender As Object, e As EventArgs) Handles txtDoubleNights.TextChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub txtKingRooms_Enter(sender As Object, e As EventArgs) Handles txtKingRooms.Enter

        txtKingRooms.SelectAll()

    End Sub

    Private Sub txtKingNights_Enter(sender As Object, e As EventArgs) Handles txtKingNights.Enter

        txtKingNights.SelectAll()

    End Sub

    Private Sub txtQueenRooms_Enter(sender As Object, e As EventArgs) Handles txtQueenRooms.Enter

        txtQueenRooms.SelectAll()

    End Sub

    Private Sub txtQueenNights_Enter(sender As Object, e As EventArgs) Handles txtQueenNights.Enter

        txtQueenNights.SelectAll()

    End Sub

    Private Sub txtDoubleRooms_Enter(sender As Object, e As EventArgs) Handles txtDoubleRooms.Enter

        txtDoubleRooms.SelectAll()

    End Sub

    Private Sub txtDoubleNights_Enter(sender As Object, e As EventArgs) Handles txtDoubleNights.Enter

        txtDoubleNights.SelectAll()

    End Sub

    Private Sub txtKingRooms_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtKingRooms.KeyPress

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtKingNights_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtKingNights.KeyPress

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtQueenRooms_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQueenRooms.KeyPress

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtQueenNights_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQueenNights.KeyPress

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtDoubleRooms_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDoubleRooms.KeyPress

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtDoubleNights_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDoubleNights.KeyPress

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        txtKingRooms.Text = String.Empty
        txtKingNights.Text = String.Empty
        txtQueenRooms.Text = String.Empty
        txtQueenNights.Text = String.Empty
        txtDoubleRooms.Text = String.Empty
        txtDoubleNights.Text = String.Empty
        txtPromo.Text = String.Empty

        chkSelectKing.Checked = False
        chkSelectQueen.Checked = False
        chkSelectDouble.Checked = False

        radMemberKing.Checked = False
        radRegularKing.Checked = False
        radMemeberQueen.Checked = False
        radRegularQueen.Checked = False
        radMemberDouble.Checked = False
        radRegularDouble.Checked = False

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty
    End Sub

    Private Sub radMemberKing_CheckedChanged(sender As Object, e As EventArgs) Handles radMemberKing.CheckedChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub radRegularKing_CheckedChanged(sender As Object, e As EventArgs) Handles radRegularKing.CheckedChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub radMemeberQueen_CheckedChanged(sender As Object, e As EventArgs) Handles radMemeberQueen.CheckedChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub radRegularQueen_CheckedChanged(sender As Object, e As EventArgs) Handles radRegularQueen.CheckedChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub radMemberDouble_CheckedChanged(sender As Object, e As EventArgs) Handles radMemberDouble.CheckedChanged

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub radRegularDouble_Click(sender As Object, e As EventArgs) Handles radRegularDouble.Click

        lblTotalRoomsBooked.Text = String.Empty
        lblTotalGuestsBooked.Text = String.Empty
        lblTotalNightsBooked.Text = String.Empty
        lblRoomCharge.Text = String.Empty
        lblTax.Text = String.Empty
        lblHotelFees.Text = String.Empty
        lblTotalDue.Text = String.Empty

    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub
End Class
