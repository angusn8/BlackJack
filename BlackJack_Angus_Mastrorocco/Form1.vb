Public Class Form1

    Dim CardNumber(52) As Integer
    Dim CardIndex As Integer
    Dim TempValue As Integer
    Dim LoopCounter As Integer
    Dim ItemPicked As Integer
    Dim Remaining As Integer
    Dim MyRandom As New Random
    Dim YourNumber As Integer 'Your Card Number
    Dim YourNumber2 As Integer
    Dim YourNumber3 As Integer
    Dim YourNumber4 As Integer
    Dim YourNumber5 As Integer
    Dim ComputerNumber As Integer 'Computer's Card Number
    Dim ComputerNumber2 As Integer
    Dim ComputerNumber3 As Integer
    Dim ComputerNumber4 As Integer
    Dim ComputerNumber5 As Integer
    Dim PlayerScore As Integer
    Dim DealScore As Integer
    Dim PlayerGame As Integer
    Dim DealGame As Integer

    Private Sub btnNewGame_Click(sender As Object, e As EventArgs) Handles btnNewGame.Click

        'New Game is CLicked
        If btnNewGame.Text = "New Game" Then

            btnHit.Visible = True
            btnStay.Visible = True

            'Change Scores to Zero 
            PlayerScore = 0
            txtPlayerScore.Text = PlayerScore
            DealScore = 0
            txtDealScore.Text = DealScore

            'Shuffle Cards
            For LoopCounter = 1 To 52
                CardNumber(LoopCounter) = LoopCounter
            Next LoopCounter

            'Start at 52 and swap one value 
            For Remaining = 52 To 2 Step -1
                'Pick random item 
                ItemPicked = MyRandom.Next(Remaining) + 1
                'Swap Picked with bottom 
                TempValue = CardNumber(Remaining)
                CardNumber(Remaining) = CardNumber(ItemPicked)
                CardNumber(ItemPicked) = TempValue
            Next Remaining
            'Set card index to 1 
            CardIndex = 1
        End If

        'Make Cards visible 
        pnlCard1.Visible = True
        pnlCard2.Visible = True
        pnlCard3.Visible = False
        pnlCard4.Visible = False
        pnlCard5.Visible = False
        pnlCard6.Visible = True
        pnlCard7.Visible = False
        pnlCard8.Visible = False
        pnlCard9.Visible = False
        pnlCard10.Visible = False

        'Result textbox cleared 
        txtResult.Text = " "

        'Display users first Card Suit and Number 
        Select Case CardNumber(1)
            Case 1 To 13
                picCard1.Image = picHeart.Image
                YourNumber = CardNumber(1)
            Case 14 To 26
                picCard1.Image = picDiamond.Image
                YourNumber = CardNumber(1) - 13
            Case 27 To 39
                picCard1.Image = picClub.Image
                YourNumber = CardNumber(1) - 26
            Case 40 To 52
                picCard1.Image = picSpade.Image
                YourNumber = CardNumber(1) - 39
        End Select

        'Display card value
        Select Case YourNumber
            Case 1 To 9
                lblCard1.Text = Str(YourNumber + 1) + " "
                YourNumber = YourNumber + 1
            Case 10
                lblCard1.Text = "J"
            Case 11
                lblCard1.Text = "Q"
                YourNumber = 10
            Case 12
                lblCard1.Text = "K"
                YourNumber = 10
            Case 13
                lblCard1.Text = "A"
                YourNumber = 11
        End Select

        'Display users Second Card Suit and Number 
        Select Case CardNumber(2)
            Case 1 To 13
                picCard2.Image = picHeart.Image
                YourNumber2 = CardNumber(2)
            Case 14 To 26
                picCard2.Image = picDiamond.Image
                YourNumber2 = CardNumber(2) - 13
            Case 27 To 39
                picCard2.Image = picClub.Image
                YourNumber2 = CardNumber(2) - 26
            Case 40 To 52
                picCard2.Image = picSpade.Image
                YourNumber2 = CardNumber(2) - 39
        End Select

        'Display card value
        Select Case YourNumber2
            Case 1 To 9
                lblCard2.Text = Str(YourNumber2 + 1) + " "
                YourNumber2 = YourNumber2 + 1
            Case 10
                lblCard2.Text = "J"
            Case 11
                lblCard2.Text = "Q"
                YourNumber2 = 10
            Case 12
                lblCard2.Text = "K"
                YourNumber2 = 10
            Case 13
                lblCard2.Text = "A"
                YourNumber2 = 11
        End Select

        'Display users third Card Suit and Number 
        Select Case CardNumber(3)
            Case 1 To 13
                picCard3.Image = picHeart.Image
                YourNumber3 = CardNumber(3)
            Case 14 To 26
                picCard3.Image = picDiamond.Image
                YourNumber3 = CardNumber(3) - 13
            Case 27 To 39
                picCard3.Image = picClub.Image
                YourNumber3 = CardNumber(3) - 26
            Case 40 To 52
                picCard3.Image = picSpade.Image
                YourNumber3 = CardNumber(3) - 39
        End Select

        'Display card value
        Select Case YourNumber3
            Case 1 To 9
                lblCard3.Text = Str(YourNumber3 + 1) + " "
                YourNumber3 = YourNumber3 + 1
            Case 10
                lblCard3.Text = "J"
            Case 11
                lblCard3.Text = "Q"
                YourNumber3 = 10
            Case 12
                lblCard3.Text = "K"
                YourNumber3 = 10
            Case 13
                lblCard3.Text = "A"
                YourNumber3 = 11
        End Select

        'Display users fourth Card Suit and Number 
        Select Case CardNumber(4)
            Case 1 To 13
                picCard4.Image = picHeart.Image
                YourNumber4 = CardNumber(4)
            Case 14 To 26
                picCard4.Image = picDiamond.Image
                YourNumber4 = CardNumber(4) - 13
            Case 27 To 39
                picCard4.Image = picClub.Image
                YourNumber4 = CardNumber(4) - 26
            Case 40 To 52
                picCard4.Image = picSpade.Image
                YourNumber4 = CardNumber(4) - 39
        End Select

        'Display card value
        Select Case YourNumber4
            Case 1 To 9
                lblCard4.Text = Str(YourNumber4 + 1) + " "
                YourNumber4 = YourNumber4 + 1
            Case 10
                lblCard4.Text = "J"
            Case 11
                lblCard4.Text = "Q"
                YourNumber4 = 10
            Case 12
                lblCard4.Text = "K"
                YourNumber4 = 10
            Case 13
                lblCard4.Text = "A"
                YourNumber4 = 11
        End Select

        'Display users fifth Card Suit and Number 
        Select Case CardNumber(5)
            Case 1 To 13
                picCard5.Image = picHeart.Image
                YourNumber5 = CardNumber(5)
            Case 14 To 26
                picCard5.Image = picDiamond.Image
                YourNumber5 = CardNumber(5) - 13
            Case 27 To 39
                picCard5.Image = picClub.Image
                YourNumber5 = CardNumber(5) - 26
            Case 40 To 52
                picCard5.Image = picSpade.Image
                YourNumber5 = CardNumber(5) - 39
        End Select

        'Display card value
        Select Case YourNumber5
            Case 1 To 9
                lblCard5.Text = Str(YourNumber5 + 1) + " "
                YourNumber5 = YourNumber5 + 1
            Case 10
                lblCard5.Text = "J"
            Case 11
                lblCard5.Text = "Q"
                YourNumber5 = 10
            Case 12
                lblCard5.Text = "K"
                YourNumber5 = 10
            Case 13
                lblCard5.Text = "A"
                YourNumber5 = 11
        End Select

        'Display computer's first card
        Select Case CardNumber(6)
            Case 1 To 13
                picCard6.Image = picHeart.Image
                ComputerNumber = CardNumber(6)
            Case 14 To 26
                picCard6.Image = picDiamond.Image
                ComputerNumber = CardNumber(6) - 13
            Case 27 To 39
                picCard6.Image = picClub.Image
                ComputerNumber = CardNumber(6) - 26
            Case 40 To 52
                picCard6.Image = picSpade.Image
                ComputerNumber = CardNumber(6) - 39
        End Select

        'display card value
        Select Case ComputerNumber
            Case 1 To 9
                lblCard6.Text = Str(ComputerNumber + 1) + " "
                ComputerNumber = ComputerNumber + 1
            Case 10
                lblCard6.Text = "J"
            Case 11
                lblCard6.Text = "Q"
                ComputerNumber = 10
            Case 12
                lblCard6.Text = "K"
                ComputerNumber = 10
            Case 13
                lblCard6.Text = "A"
                ComputerNumber = 11
        End Select

        'Display computer's second card
        Select Case CardNumber(7)
            Case 1 To 13
                picCard7.Image = picHeart.Image
                ComputerNumber2 = CardNumber(7)
            Case 14 To 26
                picCard7.Image = picDiamond.Image
                ComputerNumber2 = CardNumber(7) - 13
            Case 27 To 39
                picCard7.Image = picClub.Image
                ComputerNumber2 = CardNumber(7) - 26
            Case 40 To 52
                picCard7.Image = picSpade.Image
                ComputerNumber2 = CardNumber(7) - 39
        End Select

        'display card value
        Select Case ComputerNumber2
            Case 1 To 9
                lblCard7.Text = Str(ComputerNumber2 + 1) + " "
                ComputerNumber2 = ComputerNumber2 + 1
            Case 10
                lblCard7.Text = "J"
            Case 11
                lblCard7.Text = "Q"
                ComputerNumber2 = 10
            Case 12
                lblCard7.Text = "K"
                ComputerNumber2 = 10
            Case 13
                lblCard7.Text = "A"
                ComputerNumber2 = 11
        End Select

        'Display computer's third card
        Select Case CardNumber(8)
            Case 1 To 13
                picCard8.Image = picHeart.Image
                ComputerNumber3 = CardNumber(8)
            Case 14 To 26
                picCard8.Image = picDiamond.Image
                ComputerNumber3 = CardNumber(8) - 13
            Case 27 To 39
                picCard8.Image = picClub.Image
                ComputerNumber3 = CardNumber(8) - 26
            Case 40 To 52
                picCard8.Image = picSpade.Image
                ComputerNumber3 = CardNumber(8) - 39
        End Select

        'display card value
        Select Case ComputerNumber3
            Case 1 To 9
                lblCard8.Text = Str(ComputerNumber3 + 1) + " "
                ComputerNumber3 = ComputerNumber3 + 1
            Case 10
                lblCard8.Text = "J"
            Case 11
                lblCard8.Text = "Q"
                ComputerNumber3 = 10
            Case 12
                lblCard8.Text = "K"
                ComputerNumber3 = 10
            Case 13
                lblCard8.Text = "A"
                ComputerNumber3 = 11
        End Select

        'Display computer's fourth card
        Select Case CardNumber(9)
            Case 1 To 13
                picCard9.Image = picHeart.Image
                ComputerNumber4 = CardNumber(9)
            Case 14 To 26
                picCard9.Image = picDiamond.Image
                ComputerNumber4 = CardNumber(9) - 13
            Case 27 To 39
                picCard9.Image = picClub.Image
                ComputerNumber4 = CardNumber(9) - 26
            Case 40 To 52
                picCard9.Image = picSpade.Image
                ComputerNumber4 = CardNumber(9) - 39
        End Select

        'display card value
        Select Case ComputerNumber4
            Case 1 To 9
                lblCard9.Text = Str(ComputerNumber4 + 1) + " "
                ComputerNumber4 = ComputerNumber4 + 1
            Case 10
                lblCard9.Text = "J"
            Case 11
                lblCard9.Text = "Q"
                ComputerNumber4 = 10
            Case 12
                lblCard9.Text = "K"
                ComputerNumber4 = 10
            Case 13
                lblCard9.Text = "A"
                ComputerNumber4 = 11
        End Select

        'Display computer's fifth card
        Select Case CardNumber(10)
            Case 1 To 13
                picCard10.Image = picHeart.Image
                ComputerNumber5 = CardNumber(10)
            Case 14 To 26
                picCard10.Image = picDiamond.Image
                ComputerNumber5 = CardNumber(10) - 13
            Case 27 To 39
                picCard10.Image = picClub.Image
                ComputerNumber5 = CardNumber(10) - 26
            Case 40 To 52
                picCard10.Image = picSpade.Image
                ComputerNumber5 = CardNumber(10) - 39
        End Select

        'display card value
        Select Case ComputerNumber5
            Case 1 To 9
                lblCard10.Text = Str(ComputerNumber5 + 1) + " "
                ComputerNumber5 = ComputerNumber5 + 1
            Case 10
                lblCard10.Text = "J"
            Case 11
                lblCard10.Text = "Q"
                ComputerNumber5 = 10
            Case 12
                lblCard10.Text = "K"
                ComputerNumber5 = 10
            Case 13
                lblCard10.Text = "A"
                ComputerNumber5 = 11
        End Select

        'Calculate score 
        PlayerScore = YourNumber + YourNumber2
        DealScore = ComputerNumber
        txtPlayerScore.Text = Str(PlayerScore)
        txtDealScore.Text = Str(DealScore)

        'Add to player game score if blackjack 
        If YourNumber + YourNumber2 = 21 Then
            PlayerGame = PlayerGame + 1
            txtPlayerGame.Text = Str(PlayerGame)
        End If

        'Check Blackjack 
        If PlayerScore = 21 Then
            txtResult.Text = "BlackJack"
            btnHit.Visible = False
            btnStay.Visible = False
        End If
    End Sub

    Private Sub btnHit_Click(sender As Object, e As EventArgs) Handles btnHit.Click

        'Display player's cards in order
        If pnlCard3.Visible = False Then
            pnlCard3.Visible = True
            PlayerScore = YourNumber + YourNumber2 + YourNumber3
            txtPlayerScore.Text = Str(PlayerScore)
        ElseIf pnlCard4.Visible = False Then
            pnlCard4.Visible = True
            PlayerScore = YourNumber + YourNumber2 + YourNumber3 + YourNumber4
            txtPlayerScore.Text = Str(PlayerScore)
        ElseIf pnlCard5.Visible = False Then
            pnlCard5.Visible = True
        PlayerScore = YourNumber + YourNumber2 + YourNumber3 + YourNumber4 + YourNumber5
        txtPlayerScore.Text = Str(PlayerScore)
        End If

        ' Check for bust 
        If PlayerScore > 21 Then
            txtResult.Text = "Bust"
            btnHit.Visible = False
            btnStay.Visible = False
            DealGame = DealGame + 1
            txtDealGame.Text = Str(DealGame)
        End If

        'If 5 cards, Then win 
        If pnlCard5.Visible = True And PlayerScore <= 21 Then
            txtResult.Text = "Win"
            btnHit.Visible = False
            btnStay.Visible = False
        End If

    End Sub

    Private Sub btnStay_Click(sender As Object, e As EventArgs) Handles btnStay.Click

        'Diplay dealer's second card when player turn ends and calculate score
        pnlCard7.Visible = True
        DealScore = ComputerNumber + ComputerNumber2
        txtDealScore.Text = Str(DealScore)

        'Decide how many more cards to display based on score 
        If pnlCard8.Visible = False And DealScore < 17 Then
            pnlCard8.Visible = True
            DealScore = ComputerNumber + ComputerNumber2 + ComputerNumber3
            txtDealScore.Text = Str(DealScore)
        ElseIf pnlCard9.Visible = False And DealScore < 17 Then
            pnlCard9.Visible = True
            DealScore = ComputerNumber + ComputerNumber2 + ComputerNumber3 + ComputerNumber4
            txtDealScore.Text = Str(DealScore)
        ElseIf pnlCard10.Visible = False And DealScore < 17 Then
            pnlCard10.Visible = True
            DealScore = ComputerNumber + ComputerNumber2 + ComputerNumber3 + ComputerNumber4 + ComputerNumber5
            txtDealScore.Text = Str(DealScore)
        End If

        'Check for bust 
        If DealScore > 21 Then
            txtResult.Text = "Dealer Bust"
            btnHit.Visible = False
            btnStay.Visible = False
        End If

        'Check winner and display 
        If PlayerScore > DealScore Then
            txtResult.Text = "You Win"
            PlayerGame = PlayerGame + 1
            btnHit.Visible = False
            btnStay.Visible = False
            txtPlayerGame.Text = Str(PlayerGame)
        ElseIf DealScore > PlayerScore And DealScore > 21 Then
            txtResult.Text = "Dealer Bust"
            btnHit.Visible = False
            btnStay.Visible = False
            PlayerGame = PlayerGame + 1
            txtPlayerGame.Text = Str(PlayerGame)
        ElseIf DealScore > PlayerScore Then
            txtResult.Text = "Dealer Wins"
            DealGame = DealGame + 1
            btnHit.Visible = False
            btnStay.Visible = False
            txtDealGame.Text = Str(DealGame)
        ElseIf DealScore = PlayerScore Then
            txtResult.Text = "Push"
            btnHit.Visible = False
            btnStay.Visible = False
        End If

    End Sub

    Private Sub btnDirections_Click(sender As Object, e As EventArgs) Handles btnDirections.Click

        'Display message box with direction 
        MessageBox.Show("1.) You play against the computer to see who gets the closest to '21'
2.) Cards are worth their value (2=2, 3=3, ect,), Face cars (J,Q,K) are worth 10, and Ace's are worth 11
3.) If you go over 21 you lose automatically (Bust)
4.) You start with 2 cards, the computer starts with two cards (one hidden one revealed.)
5.) You can choose the 'Hit' and get a new card, or 'Stay' if over 17.
6.) The computer must 'Hit' if under 17 and must 'Stay' if over 17.
7.) If you accumulate 5 cards without going over 21 you win automatically
8.) A tie is called a 'Push'
9.) Blackjack wins automatically. Blackjack is any Ace with another facecard.")


    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        'Close the form
        Me.Close()

    End Sub

End Class