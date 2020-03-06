Public Class Form1
    Structure Card

        Dim V As Integer
        Dim S As Integer


        Function getcstr() As String
            getcstr = ""
            Select Case V
                Case 1
                    getcstr = getcstr + "a"

                Case 10
                    getcstr = getcstr + "t"
                Case 11
                    getcstr = getcstr + "j"
                Case 12
                    getcstr = getcstr + "q"
                Case 13
                    getcstr = getcstr + "k"
                Case Else
                    getcstr = getcstr + Trim(Str(V))
            End Select

            Select Case S
                Case 1
                    getcstr = getcstr + "c"
                Case 2
                    getcstr = getcstr + "h"
                Case 3
                    getcstr = getcstr + "d"
                Case 4
                    getcstr = getcstr + "s"
            End Select




        End Function
    End Structure
    Structure HS
        Dim fscore As Integer
        Dim fname As String

        Function geths() As String
            geths = ""
            Select Case fscore
                Case 1
                    geths = 0
                Case Else
                    geths = fscore
            End Select

            Select Case fname
                        Case 1

                        Case Else
                    geths = geths + "  " + Trim(fname)
            End Select
        End Function



    End Structure

    Function Shuffle()
        For i = 1 To 4
            For j = 1 To 13
                C = C + 1
                Deck(C).S = i
                Deck(C).V = j
            Next j
        Next i

        Randomize()

        For i = 1 To 52
            rnum = Int(Rnd() * 52) + 1
            temp = Deck(i)
            Deck(i) = Deck(rnum)
            Deck(rnum) = temp
        Next i
    End Function
    Dim Deck(52) As Card
    Dim temp As Card
    Dim counter As Integer, dcounter As Integer
    Dim arrpic(5) As PictureBox
    Dim arrdpic(5) As PictureBox
    Dim rnum As Integer, rnum1 As Integer
    Dim j As Integer
    Dim C As Integer
    Dim i As Integer
    Dim num As Integer
    Dim dnum As Integer
    Dim tempnum As Integer
    Dim ans As String
    Dim filename As String
    Dim bob As String
    Dim path As String
    Dim score As Integer
    Dim bet As Integer
    Dim aflag As Integer
    Dim cardcounter As Integer
    Dim strline As String
    Dim arrhs(11) As HS
    Dim rarrhs(11) As String
    Dim arrfscores(11) As Integer
    Dim arrfnames(11) As String
    Dim temphs As HS
    Dim fscore As String
    Dim frontnum As Integer
    Dim backnum As Integer
    Public AceValue As Integer
    Dim dtempnum As Integer
    Dim tempscore As Integer
    Public insurance As Boolean

    Sub Picarray()
        arrpic(1) = pic1
        arrpic(2) = pic2
        arrpic(3) = pic3
        arrpic(4) = pic4
        arrpic(5) = pic5
        arrdpic(1) = DPic1
        arrdpic(2) = Dpic2
        arrdpic(3) = Dpic3
        arrdpic(4) = dpic4
        arrdpic(5) = dpic5
    End Sub
    Private Sub cmdHit_Click(sender As Object, e As EventArgs) Handles cmdHit.Click
        Button3.Visible = False
        rnum = Int(Rnd() * 52) + 1

        If counter = 1 Then
            pic1.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
            If Deck(rnum).V >= 10 Then
                num = num + 10
                tempnum = tempnum + 10
            ElseIf Deck(rnum).V = 1 Then
                Dialog1.ShowDialog()
                num = num + AceValue
                TextBox1.Text = num
                tempnum = tempnum + 1
                TextBox2.Visible = True
                TextBox2.Text = tempnum
                aflag = 2
            Else
                num = num + Deck(rnum).V
                tempnum = tempnum + Deck(rnum).V
            End If
            Label1.Text = Str(num)
            Label2.Text = tempnum

        ElseIf counter = 2 Then
            pic2.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
            pic2.Visible = True
            If Deck(rnum).V >= 10 Then
                num = num + 10
                tempnum = tempnum + 10
            ElseIf Deck(rnum).V = 1 Then
                Dialog1.ShowDialog()
                num = num + AceValue
                TextBox1.Text = num
                tempnum = tempnum + 1
                TextBox2.Visible = True
                TextBox2.Text = tempnum
                aflag = 2
            Else
                num = num + Deck(rnum).V
                tempnum = tempnum + Deck(rnum).V
            End If

            Label1.Text = num
            Label2.Text = tempnum
        ElseIf counter = 3 Then
            pic3.Location = New Point(0, 0)
            pic3.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
            pic3.Visible = True
            If Deck(rnum).V >= 10 Then
                num = num + 10
                tempnum = tempnum + 10
            ElseIf Deck(rnum).V = 1 Then
                Dialog1.ShowDialog()
                num = num + AceValue
                TextBox1.Text = num
                tempnum = tempnum + 1
                TextBox2.Visible = True
                TextBox2.Text = tempnum
                aflag = 2
            Else
                num = num + Deck(rnum).V
                tempnum = tempnum + Deck(rnum).V
            End If

            Label1.Text = num
            Label2.Text = tempnum
            Timer3.Enabled = True
        ElseIf counter = 4 Then
            pic4.Location = New Point(0, 0)
            Timer4.Enabled = True
            pic4.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
            pic4.Visible = True
            If Deck(rnum).V >= 10 Then
                num = num + 10
                tempnum = tempnum + 10
            ElseIf Deck(rnum).V = 1 Then
                Dialog1.ShowDialog()
                num = num + AceValue
                TextBox1.Text = num
                tempnum = tempnum + 1
                TextBox2.Visible = True
                TextBox2.Text = tempnum
                aflag = 2
            Else
                num = num + Deck(rnum).V
                tempnum = tempnum + Deck(rnum).V
            End If

            Label1.Text = num
            Label2.Text = tempnum

        Else
            pic5.Location = New Point(0, 0)
            Timer5.Enabled = True
            pic5.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
            pic5.Visible = True
            If Deck(rnum).V >= 10 Then
                num = num + 10
                tempnum = tempnum + 10
            ElseIf Deck(rnum).V = 1 Then
                Dialog1.Show()
                num = num + AceValue
                TextBox1.Text = num
                tempnum = tempnum + 1
                TextBox2.Visible = True
                TextBox2.Text = tempnum
                aflag = 2
            Else
                num = num + Deck(rnum).V
                tempnum = tempnum + Deck(rnum).V
            End If

            If num <= 21 Or tempnum <= 21 Then
                TextBox1.Text = "You Won due to Five Card Charlie"
                score = score - bet
                score = score + bet
                score = score + bet
                lblScore.Text = score
            End If
            Label1.Text = num
            Label2.Text = tempnum

        End If

        If tempnum = num Then
            If num > 21 Then
                TextBox1.Text = "Dealer Won"
                score = score - bet
                lblScore.Text = score

                'ElseIf num < 21 And pic5.Visible = True Then
                'TextBox1.Text = "You Won"
                'score = score + bet
                'lblScore.Text = score
            ElseIf num = 21 Then


            ElseIf num > 21 Then
                TextBox1.Text = "Dealer Won"
                score = score - bet
                lblScore.Text = score
                'TextBox1.Text = "You Won"
                'score = score + bet
                'lblScore.Text = score
                'Message box appears play again and another for bet
            End If

        Else
            If num > 21 And tempnum < 21 Then
                TextBox1.Text = "Ace Value Changed"
            ElseIf num > 21 And tempnum > 21 Then
                TextBox1.Text = "Dealer Won"
                score = score - bet
                lblScore.Text = score
            ElseIf num < 21 And tempnum < 21 Then

            End If



        End If



        counter = counter + 1
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer
        ListBoxHigh.Items.Clear()
        Dim ofile As System.IO.File
        Dim oread As System.IO.StreamReader
        oread = ofile.OpenText("C:\Data\CP2\fscores.dat")
        For i = 1 To 10
            arrhs(i).fscore = Val(oread.ReadLine())
            arrhs(i).fname = oread.ReadLine()

        Next i

        oread.Close()

        For i = 1 To 10
            ListBoxHigh.Items.Add(arrhs(i).fscore.ToString + "   " + arrhs(i).fname)
        Next



        counter = 1
        cardcounter = 0


        i = 0
        'FileOpen(1, "E:\fscores.dat", OpenMode.Input)

        'Do While Not EOF(1)
        'arrhs(i).fscore = LineInput(1)

        'ListBoxHigh.Items.Add(arrhs(i)) 'sending to a combo box
        'i = i + 1
        'Loop
        'FileClose(1)

        'Dim i As Integer
        i = 0
        'Dim i As Integer
        'i = 0
        'FileOpen(1, "E:\MVfnames.dat", OpenMode.Input)

        ' Do While Not EOF(1)
        'arrhs(i).fname = LineInput(1)

        'ListBoxHigh.Items.Add(arrhs(i)) 'sending to a combo box
        'i = i + 1
        'Loop
        'FileClose(1)




        'pichand.Location = New Point(29, -91)
        'pic1.Location = New Point(12, 12)
        'pic2.Location = New Point(12, 12)

        pic1.SendToBack()
        pic2.SendToBack()
        pic3.SendToBack()
        pic4.SendToBack()
        pic5.SendToBack()

        Shuffle()

        ans = vbNo
        Do While ans = vbNo
            filename = UCase(Trim(InputBox("Enter Username: ", "", "EV")))


            Exit Do

        Loop




        score = 100
        lblScore.Text = score

        'ans = vbNo
        ' Do While ans = vbNo
        'bet = UCase(Trim(InputBox("Enter Wager.. ", "", "0 - 100")))


        'Exit Do

        'Loop
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ListBoxHigh.Items.Clear()
        Dim ofile As System.IO.File
        Dim oread As System.IO.StreamReader
        oread = ofile.OpenText("C:\Data\CP2\fscores.dat")
        For i = 1 To 10
            arrhs(i).fscore = Val(oread.ReadLine())
            arrhs(i).fname = oread.ReadLine()

        Next i

        oread.Close()



        If score > Val(arrhs(10).fscore) Then
            arrhs(10).fscore = score
            arrhs(10).fname = filename
            For frontnum = 1 To 9
                For backnum = frontnum + 1 To 10
                    If arrhs(frontnum).fscore < arrhs(backnum).fscore Then
                        temphs = arrhs(frontnum)
                        arrhs(frontnum) = arrhs(backnum)
                        arrhs(backnum) = temphs
                    End If
                Next backnum
            Next frontnum

        End If

        Dim owrite As System.IO.StreamWriter

        owrite = ofile.CreateText("C:\Data\CP2\fscores.dat")

        For i = 1 To 10
            owrite.Write(arrhs(i).fscore.ToString)
            owrite.WriteLine()
            owrite.Write(arrhs(i).fname)
            owrite.WriteLine()
        Next i
        owrite.Close()

        For i = 1 To 10
            ListBoxHigh.Items.Add(arrhs(i).fscore.ToString + "   " + arrhs(i).fname)
        Next i



        'arrhs(11).fscore = score
        '(11).fname = filename



        'For frontnum = 1 To 10
        'For backnum = frontnum + 1 To 11
        'If arrhs(frontnum).fscore < arrhs(backnum).fscore Then
        'temphs = arrhs(frontnum)
        'arrhs(frontnum).fscore = arrhs(backnum).fscore
        'arrhs(backnum) = temphs
        'End If
        'Next backnum
        'Next frontnum

        'ListBoxHigh.Items.Clear()

        ' For i = 1 To 10
        'ListBoxHigh.Items.Add(arrhs(i).geths)
        'Next i

        'For i = 1 To 10
        'arrfscores(i) = arrhs(i).fscore
        'arrfnames(i) = arrhs(i).fname
        ' Next i
        ' path = "E:\fscores.dat"
        ' 'FileOpen(1, "E:\JohnDickinson.dat", OpenMode.Output)
        'For i = 1 To VinList.Items.Count - 1
        ' IO.File.WriteAllLines(path, arrfscores)

        'path = "E:\fnames.dat"
        'IO.File.WriteAllLines(path, arrfnames)





    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        pichand.Visible = True
        pichand.Location = New Point(pichand.Location.X, pichand.Location.Y + 2)
        pic1.Location = New Point(pic1.Location.X + 5, pic1.Location.Y + 7.5)


        'Do While pic1.Location.X + 100 > pic3.Location.X Then
        If pic1.Location.X + 100 > pic3.Location.X Then
            pic1.Location = New Point(94, 162)
            Timer1.Enabled = False
            Timer2.Enabled = True
            pichand.Location = New Point(29, -80)
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        pichand.Visible = True
        pic2.Location = New Point(pic2.Location.X + 4, pic2.Location.Y + 2)
        pichand.Visible = True
        pichand.Location = New Point(pichand.Location.X, pichand.Location.Y + 2)
        pic2.Visible = True
        If pic2.Location.Y + 75 > 162 Then
            pic2.Location = New Point(130, 162)
            'pichand.Visible = True
            pichand.Location = New Point(pichand.Location.X, pichand.Location.Y + 2)
            pichand.Visible = False
            Timer2.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        arrhs(1).fname = "EB"
        arrhs(2).fname = "CW"
        arrhs(3).fname = "MV"
        arrhs(4).fname = "WM"
        arrhs(5).fname = "GR"
        arrhs(6).fname = "YR"
        arrhs(7).fname = "KY"
        arrhs(8).fname = "BW"
        arrhs(9).fname = "PS"
        arrhs(10).fname = "KR"

        arrhs(1).fscore = 756
        arrhs(2).fscore = 385
        arrhs(3).fscore = 654
        arrhs(4).fscore = 348
        arrhs(5).fscore = 275
        arrhs(6).fscore = 450
        arrhs(7).fscore = 150
        arrhs(8).fscore = 98
        arrhs(9).fscore = 50
        arrhs(10).fscore = 1

        Dim owrite As System.IO.StreamWriter
        Dim ofile As System.IO.File
        owrite = ofile.CreateText("C:\Data\CP2\fscores.dat")

        For i = 1 To 10
            owrite.Write(arrhs(i).fscore.ToString)
            owrite.WriteLine()
            owrite.Write(arrhs(i).fname)
            owrite.WriteLine()
        Next i
        owrite.Close()

        For i = 1 To 10
            ListBoxHigh.Items.Add(arrhs(i).fscore.ToString + "   " + arrhs(i).fname)
        Next i



    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        pichand.Visible = True

        pic3.Visible = True
        pichand.Location = New Point(pichand.Location.X, pichand.Location.Y + 2)
        pic3.Location = New Point(pic3.Location.X + 7.5, pic3.Location.Y + 7.5)


        'Do While pic1.Location.X + 100 > pic3.Location.X Then
        If pic3.Location.X + 100 > pic4.Location.X Then
            pic3.Location = New Point(150, 162)
            Timer3.Enabled = False
            'Timer2.Enabled = True
            pichand.Location = New Point(29, -80)
            pichand.Visible = False
        End If
    End Sub

    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        pichand.Visible = True
        pic4.Visible = True
        pichand.Location = New Point(pichand.Location.X, pichand.Location.Y + 2)
        pic4.Location = New Point(pic4.Location.X + 10, pic4.Location.Y + 7.5)


        'Do While pic1.Location.X + 100 > pic3.Location.X Then
        If pic4.Location.X + 100 > pic5.Location.X Then
            pic4.Location = New Point(170, 162)
            Timer4.Enabled = False
            'Timer2.Enabled = True
            pichand.Location = New Point(29, -80)
            pichand.Visible = False
        End If
    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        pichand.Visible = True
        pic5.Visible = True
        pichand.Location = New Point(pichand.Location.X, pichand.Location.Y + 2)
        pic5.Location = New Point(pic5.Location.X + 7.5, pic5.Location.Y + 7.5)


        'Do While pic1.Location.X + 100 > pic3.Location.X Then
        If pic5.Location.Y + 100 > pic1.Location.Y Then
            pic5.Location = New Point(190, 162)
            Timer5.Enabled = False
            'Timer2.Enabled = True
            pichand.Location = New Point(29, -80)
            pichand.Visible = False
        End If
    End Sub

    Private Sub pic5_Click(sender As Object, e As EventArgs) Handles pic5.Click

    End Sub

    Private Sub pic4_Click(sender As Object, e As EventArgs) Handles pic4.Click

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If bet * 2 > score Then
            MsgBox("You do not have enough points!")
        Else
            bet = bet * 2


            pic3.Location = New Point(0, 0)
            pic3.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
            'pic3.Visible = True
            If Deck(rnum).V >= 10 Then
                num = num + 10
                tempnum = tempnum + 10
            ElseIf Deck(rnum).V = 1 Then
                num = num + 11
                TextBox1.Text = num
                tempnum = tempnum + 1
                TextBox2.Visible = True
                TextBox2.Text = tempnum
                aflag = 2
            Else
                num = num + Deck(rnum).V
                tempnum = tempnum + Deck(rnum).V
            End If

            Label1.Text = num
            Label2.Text = tempnum
            Timer3.Enabled = True

            dpic.Visible = False

            DPic1.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum1).getcstr()) + ".gif"
            DPic1.Visible = True
            If dnum > num Or dnum > tempnum And num > 21 Then
                TextBox1.Text = "Dealer Won!"
                score = score - bet
                lblScore.Text = score
            ElseIf dnum = num And dnum = 21 Or dnum = tempnum And num > 21 Then
                TextBox1.Text = "Tie!"
            Else

                For i = 3 To 5
                    rnum = Int(Rnd() * 52) + 1
                    If i = 3 Then
                        Dpic3.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
                        Dpic3.Visible = True

                        If Deck(rnum).V >= 10 Then
                            dnum = dnum + 10
                            dtempnum = dtempnum + 10
                        ElseIf Deck(rnum).V = 1 Then
                            dnum = dnum + 11
                            TextBox1.Text = dnum
                            dtempnum = dtempnum + 1
                            TextBox2.Visible = True
                            TextBox2.Text = dtempnum
                        Else
                            dnum = dnum + Deck(rnum).V
                            dtempnum = dtempnum + Deck(rnum).V
                        End If
                        If dnum > 21 And dtempnum > 21 Then
                            TextBox1.Text = "Dealer Won!"
                            score = score - bet


                            lblScore.Text = score
                            Exit For
                        ElseIf dnum <= 21 And dnum > num Or dtempnum >= num And dtempnum <= 21 Or dnum > tempnum And num > 21 And dnum <= 21 Then
                            TextBox1.Text = "Dealer Won"

                            score = score - bet
                            lblScore.Text = score
                            Exit For
                        ElseIf dnum = num Or dnum = tempnum And num > 21 Or dtempnum = num And dnum > 21 Or dtempnum = tempnum And dnum > 21 And num > 21 Then
                            lblScore.Text = "Tie! No loss or gain!"
                            Exit For
                        End If
                    ElseIf i = 4 Then
                        dpic4.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
                        dpic4.Visible = True
                        If Deck(rnum).V >= 10 Then
                            dnum = dnum + 10
                            dtempnum = dtempnum + 10
                        ElseIf Deck(rnum).V = 1 Then
                            dnum = dnum + 11
                            TextBox1.Text = dnum
                            dtempnum = dtempnum + 1
                            TextBox2.Visible = True
                            TextBox2.Text = dtempnum
                        Else
                            dnum = dnum + Deck(rnum).V
                            dtempnum = dtempnum + Deck(rnum).V
                        End If
                        If dnum > 21 And dtempnum > 21 Then

                            TextBox1.Text = "You Won"
                            score = score - bet
                            score = score + bet
                            score = score + bet
                            lblScore.Text = score
                            Exit For
                        ElseIf dnum <= 21 And dnum > num Or dtempnum >= num And dtempnum <= 21 Or dnum > tempnum And num > 21 And dnum <= 21 Then
                            TextBox1.Text = "Dealer Won"

                            score = score - bet
                            lblScore.Text = score
                            Exit For
                        ElseIf dnum = num Or dnum = tempnum And num > 21 Or dtempnum = num And dnum > 21 Or dtempnum = tempnum And dnum > 21 And num > 21 Then
                            lblScore.Text = "Tie! No loss or gain!"
                            Exit For
                        End If
                    ElseIf i = 5 Then

                        dpic5.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
                        dpic5.Visible = True
                        If Deck(rnum).V >= 10 Then
                            dnum = dnum + 10
                            dtempnum = dtempnum + 10
                        ElseIf Deck(rnum).V = 1 Then
                            dnum = dnum + 11
                            TextBox1.Text = dnum
                            dtempnum = dtempnum + 1
                            TextBox2.Visible = True
                            TextBox2.Text = dtempnum
                        Else
                            dnum = dnum + Deck(rnum).V
                            dtempnum = dtempnum + Deck(rnum).V
                        End If
                        If dnum > 21 And dtempnum > 21 Then
                            TextBox1.Text = "You won!"
                            score = score - bet
                            score = score + bet
                            score = score + bet
                            lblScore.Text = score
                            Exit For
                        ElseIf dnum <= 21 And dnum > num Or dtempnum >= num And dtempnum <= 21 Or dnum > tempnum And num > 21 And dnum <= 21 Then
                            TextBox1.Text = "Dealer Won"

                            score = score - bet
                            lblScore.Text = score
                            Exit For
                        ElseIf dnum = num Or dnum = tempnum And num > 21 Or dtempnum = num And dnum > 21 Or dtempnum = tempnum And dnum > 21 And num > 21 Then
                            lblScore.Text = "Tie! No loss or gain!"
                            Exit For
                        ElseIf dnum < num Or dnum < tempnum And num > 21 And dnum <= 21 Then
                            TextBox1.Text = "You won!"
                            score = score - bet
                            score = score + bet
                            score = score + bet
                            lblScore.Text = score
                            Exit For
                        End If
                    End If

                    ' If dnum > num And dnum <= 21 Then
                    'Label1.Text = "Dealer Won"
                    'DPic1.ImageLocation = "E:\" + Trim(Deck(rnum).getcstr()) + ".gif"
                    ' ElseIf num > dnum And num <= 21 Then
                    'Label1.Text = "Player Won"
                    ' If





                Next i
            End If


            If tempnum = num Then
                If num > 21 Then
                    TextBox1.Text = "Dealer Won"
                    score = score - bet
                    lblScore.Text = score

                    'ElseIf num < 21 And pic5.Visible = True Then
                    'TextBox1.Text = "You Won"
                    'score = score + bet
                    'lblScore.Text = score
                ElseIf num = 21 Then


                ElseIf num > 21 Then
                    TextBox1.Text = "Dealer Won"
                    score = score - bet
                    lblScore.Text = score
                    'TextBox1.Text = "You Won"
                    'score = score + bet
                    'lblScore.Text = score
                    'Message box appears play again and another for bet
                End If

            Else
                If num > 21 And tempnum < 21 Then
                    TextBox2.Text = "Ace Value Change"
                    Label1.Text = ""
                ElseIf num > 21 And tempnum > 21 Then
                    TextBox1.Text = "Dealer Won"
                    score = score - bet
                    lblScore.Text = score
                ElseIf num < 21 And tempnum < 21 Then

                End If
            End If







            If dnum <= 21 Then
                Label5.Text = dnum
            Else
                Label5.Text = dtempnum
            End If
        End If



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        dpic.Visible = False

        DPic1.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum1).getcstr()) + ".gif"
        DPic1.Visible = True
        If dnum > num Or dnum > tempnum And num > 21 Then
            If (insurance) Then
                MsgBox("Ohhhh, interesting game. Dealer won however you have insurance. Please keep playing.", MsgBoxStyle.OkOnly)
                Do Until (MsgBoxResult.Ok)
                    i += 1
                Loop
            Else
                TextBox1.Text = "Dealer Won!"
            End If
        ElseIf dnum = num And dnum = 21 Or dnum = tempnum And num > 21 Then
            TextBox1.Text = "Tie!"
        Else

            For i = 3 To 5
                rnum = Int(Rnd() * 52) + 1
                If i = 3 Then
                    Dpic3.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
                    Dpic3.Visible = True

                    If Deck(rnum).V >= 10 Then
                        dnum = dnum + 10
                        dtempnum = dtempnum + 10
                    ElseIf Deck(rnum).V = 1 Then
                        dnum = dnum + 11
                        TextBox1.Text = dnum
                        dtempnum = dtempnum + 1
                        TextBox2.Visible = True
                        TextBox2.Text = dtempnum
                    Else
                        dnum = dnum + Deck(rnum).V
                        dtempnum = dtempnum + Deck(rnum).V
                    End If
                    If dnum > 21 And dtempnum > 21 Then
                        TextBox1.Text = "You Won!"
                        score = score - bet
                        score = score + bet
                        score = score + bet
                        lblScore.Text = score
                        Exit For
                    ElseIf dnum <= 21 And dnum > num Or dtempnum >= num And dtempnum <= 21 Or dnum > tempnum And num > 21 And dnum <= 21 Then
                        If (insurance) Then
                            MsgBox("Ohhhh, interesting game. Dealer won however you have insurance. Please keep playing.", MsgBoxStyle.OkOnly)
                            Do Until (MsgBoxResult.Ok)
                                i += 1
                            Loop
                            Exit For
                        Else
                            TextBox1.Text = "Dealer Won"

                            score = score - bet
                            lblScore.Text = score
                            Exit For
                        End If

                    ElseIf dnum = num Or dnum = tempnum And num > 21 Or dtempnum = num And dnum > 21 Or dtempnum = tempnum And dnum > 21 And num > 21 Then
                        lblScore.Text = "Tie! No loss or gain!"
                        Exit For
                    End If
                ElseIf i = 4 Then
                    dpic4.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
                    dpic4.Visible = True
                    If Deck(rnum).V >= 10 Then
                        dnum = dnum + 10
                        dtempnum = dtempnum + 10
                    ElseIf Deck(rnum).V = 1 Then
                        dnum = dnum + 11
                        TextBox1.Text = dnum
                        dtempnum = dtempnum + 1
                        TextBox2.Visible = True
                        TextBox2.Text = dtempnum
                    Else
                        dnum = dnum + Deck(rnum).V
                        dtempnum = dtempnum + Deck(rnum).V
                    End If
                    If dnum > 21 And dtempnum > 21 Then

                        TextBox1.Text = "You Won"
                        score = score - bet
                        score = score + bet
                        score = score + bet
                        lblScore.Text = score
                        Exit For
                    ElseIf dnum <= 21 And dnum > num Or dtempnum >= num And dtempnum <= 21 Or dnum > tempnum And num > 21 And dnum <= 21 Then
                        TextBox1.Text = "Dealer Won"

                        score = score - bet
                        lblScore.Text = score
                        Exit For
                    ElseIf dnum = num Or dnum = tempnum And num > 21 Or dtempnum = num And dnum > 21 Or dtempnum = tempnum And dnum > 21 And num > 21 Then
                        lblScore.Text = "Tie! No loss or gain!"
                        Exit For
                    End If
                ElseIf i = 5 Then

                    dpic5.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
                    dpic5.Visible = True
                    If Deck(rnum).V >= 10 Then
                        dnum = dnum + 10
                        dtempnum = dtempnum + 10
                    ElseIf Deck(rnum).V = 1 Then
                        dnum = dnum + 11
                        TextBox1.Text = dnum
                        dtempnum = dtempnum + 1
                        TextBox2.Visible = True
                        TextBox2.Text = dtempnum
                    Else
                        dnum = dnum + Deck(rnum).V
                        dtempnum = dtempnum + Deck(rnum).V
                    End If
                    If dnum > 21 And dtempnum > 21 Then
                        TextBox1.Text = "You won!"
                        score = score - bet
                        score = score + bet
                        score = score + bet
                        lblScore.Text = score
                        Exit For
                    ElseIf dnum <= 21 And dnum > num Or dtempnum >= num And dtempnum <= 21 Or dnum > tempnum And num > 21 And dnum <= 21 Then
                        TextBox1.Text = "Dealer Won"

                        score = score - bet
                        lblScore.Text = score
                        Exit For
                    ElseIf dnum = num Or dnum = tempnum And num > 21 Or dtempnum = num And dnum > 21 Or dtempnum = tempnum And dnum > 21 And num > 21 Then
                        lblScore.Text = "Tie! No loss or gain!"
                        Exit For
                    ElseIf dnum < num Or dnum < tempnum And num > 21 And dnum <= 21 Then
                        TextBox1.Text = "You won!"
                        score = score - bet
                        score = score + bet
                        score = score + bet
                        lblScore.Text = score
                        Exit For
                    End If
                End If

                ' If dnum > num And dnum <= 21 Then
                'Label1.Text = "Dealer Won"
                'DPic1.ImageLocation = "E:\" + Trim(Deck(rnum).getcstr()) + ".gif"
                ' ElseIf num > dnum And num <= 21 Then
                'Label1.Text = "Player Won"
                ' If





            Next i
        End If

        If dnum <= 21 Then
            Label5.Text = dnum
        Else
            Label5.Text = dtempnum
        End If
    End Sub

    Private Sub cmdDeal_Click(sender As Object, e As EventArgs) Handles cmdDeal.Click


        Button3.Visible = True
        counter = 3
        DPic1.Visible = False
        Dpic2.Visible = False
        Dpic3.Visible = False
        dpic4.Visible = False
        dpic5.Visible = False
        pic1.Visible = False
        pic2.Visible = False
        pic3.Visible = False
        pic4.Visible = False
        pic5.Visible = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        Label1.Text = ""
        Label2.Text = ""
        Label5.Text = ""


        ans = vbNo
        Do While ans = vbNo
            bet = UCase(Trim(InputBox("Enter Bet... ", "", "")))

            If bet > score Then
                MsgBox("You do not have enough to wager")
                bet = UCase(Trim(InputBox("Enter Wager... ", "", "")))
                Exit Do
            Else
                Exit Do
            End If



        Loop

        aflag = 1
        dcounter = 3
        dnum = 0
        dtempnum = 0
        num = 0
        tempnum = 0

        'DPic1.Visible = True
        dpic.Visible = True
        Timer1.Enabled = True


        rnum = Int(Rnd() * 52) + 1
        pic1.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
        cardcounter = cardcounter + 1
        pic1.Visible = True
        If Deck(rnum).V >= 10 Then
            num = num + 10
            tempnum = tempnum + 10
        ElseIf Deck(rnum).V = 1 Then
            num = num + 11
            Label1.Text = num
            tempnum = tempnum + 1
            TextBox2.Visible = True
            Label2.Text = tempnum
            aflag = 2

        Else
            num = num + Deck(rnum).V
            tempnum = tempnum + Deck(rnum).V


        End If




        rnum = Int(Rnd() * 52) + 1
        pic2.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
        'pic2.Visible = True
        If Deck(rnum).V >= 10 Then
            num = num + 10
            tempnum = tempnum + 10
        ElseIf Deck(rnum).V = 1 Then
            num = num + 11
            TextBox1.Text = num
            tempnum = tempnum + 1
            TextBox2.Visible = True
            TextBox2.Text = tempnum
            aflag = 2
        Else
            num = num + Deck(rnum).V
            tempnum = tempnum + Deck(rnum).V
        End If

        If tempnum = 2 Then
            num = 12
            tempnum = 12

        End If

        Label1.Text = num
        Label2.Text = tempnum


        rnum1 = Int(Rnd() * 52) + 1
        If Deck(rnum1).V >= 10 Then
            dnum = dnum + 10
            dtempnum = dtempnum + 10
        ElseIf Deck(rnum1).V = 1 Then
            dnum = dnum + 11
            TextBox1.Text = dnum
            dtempnum = dtempnum + 1
            TextBox2.Visible = True
            TextBox2.Text = dtempnum
        Else
            dnum = dnum + Deck(rnum1).V
            dtempnum = dtempnum + Deck(rnum1).V
        End If




        'dpic.ImageLocation = "E:\" + Trim(Deck(rnum).getcstr()) + ".gif"


        rnum = Int(Rnd() * 52) + 1
        Dpic2.ImageLocation = "C:\Data\CP2\Black_Jack-V2\Card Images\" + Trim(Deck(rnum).getcstr()) + ".gif"
        If Deck(rnum).V >= 10 Then
            dnum = dnum + 10
            dtempnum = dtempnum + 10
        ElseIf Deck(rnum).V = 1 Then
            dnum = dnum + 11
            TextBox1.Text = dnum
            dtempnum = dtempnum + 1
            TextBox2.Visible = True
            TextBox2.Text = dtempnum
            MsgBox("Would you like to purchase insurance", MsgBoxStyle.YesNo)
            If (MsgBoxResult.Yes) Then
                insurance = True
            Else
                insurance = False
            End If
            Debug.Print(insurance)
        Else
            dnum = dnum + Deck(rnum).V
            dtempnum = dtempnum + Deck(rnum).V
        End If
        Dpic2.Visible = True

        If dtempnum = 2 Then
            dnum = 12
            dtempnum = 12
        End If
        'dnum = dnum + Deck(rnum).V

        'If num = 21 Or tempnum = 21 Then
        'TextBox1.Text = "You Won"
        'End If
        tempscore = score - bet
        lblScore.Text = tempscore
    End Sub


End Class
