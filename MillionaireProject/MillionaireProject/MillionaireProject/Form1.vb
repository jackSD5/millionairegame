Imports System.Threading
Imports System.Diagnostics
Imports System.Timers
Public Class Form1

    'Debug
    Dim intDebugLoad As Integer = 1


    'Game state variables
    Dim intMoneyHeld As Integer = 0
    Dim intQuestionDifficulty As Integer = 0
    Dim intUsedQuestions()
    Dim bln5050Available As Boolean = True
    Dim blnPollAudienceAvailable As Boolean = True
    Dim blnPhoneFriendAvailable As Boolean = True
    Dim intCorrectAnswer As Integer = 0
    Dim intAnswerSelected As Integer = 0
    Dim blnCorrectAnswerSelected = False
    Dim blnWaitingForInput As Boolean = False



    Public Sub CurrentQuestion(ID As Integer)

        'Database connection 
        Dim con As New OleDb.OleDbConnection

        Dim dbProvider As String
        Dim dbSource As String
        Dim MyDocumentsFolder As String
        Dim TheDatabase As String
        Dim FullDatabasePath As String

        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim sql As String

        dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
        TheDatabase = "/GameQuestions.accdb"
        MyDocumentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        FullDatabasePath = MyDocumentsFolder & TheDatabase

        dbSource = "Data Source = " & FullDatabasePath

        con.ConnectionString = dbProvider & dbSource

        con.Open()
        sql = "SELECT * FROM Questions"

        da = New OleDb.OleDbDataAdapter(sql, con)

        da.Fill(ds, "Questions")

        ' Create a DataView to sort the rows by ID in ascending order
        Dim dv As New DataView(ds.Tables("Questions"))
        dv.Sort = "ID ASC"  ' Sort by ID in ascending order



        Dim currentRow As DataRowView = Nothing
        For Each row As DataRowView In dv
            If row("ID") = ID Then
                currentRow = row
                Exit For
            End If
        Next


        Dim CQID As String = currentRow("ID").ToString()
        Dim CQCategory As String = currentRow("Category").ToString()
        Dim CQQuestionText As String = currentRow("Question text").ToString()
        Dim CQCorrectAnswer As String = currentRow("Correct answer").ToString()
        Dim CQWrongAnswer1 As String = currentRow("Wrong answer 1").ToString()
        Dim CQWrongAnswer2 As String = currentRow("Wrong answer 2").ToString()
        Dim CQWrongAnswer3 As String = currentRow("Wrong answer 3").ToString()
        Dim CQDifficulty As String = currentRow("Question difficulty").ToString()
        Dim CQPollResults As String = currentRow("Poll results").ToString()
        Dim CQ5050Option As String = currentRow("50/50 Option").ToString()
        Dim CQExpertResponse As String = currentRow("Expert response").ToString()

        debugID.Text = CQID
        debugCategory.Text = CQCategory
        lblQuestionText.Text = CQQuestionText
        debugAnswer.Text = CQCorrectAnswer
        debugWrong1.Text = CQWrongAnswer1
        debugWrong2.Text = CQWrongAnswer2
        debugWrong3.Text = CQWrongAnswer3
        debugDifficulty.Text = CQDifficulty
        debugPoll.Text = CQPollResults
        debug5050.Text = CQ5050Option
        debugHint.Text = CQExpertResponse


    End Sub

    Private Sub btnDebugBACK_Click(sender As Object, e As EventArgs) Handles btnDebugBACK.Click
        If intDebugLoad >= 2 Then
            intDebugLoad -= 1
            btnDebugLoad.Text = ("Load question " + intDebugLoad.ToString)
        End If
    End Sub

    Private Sub btnDebugFWD_Click(sender As Object, e As EventArgs) Handles btnDebugFWD.Click
        If intDebugLoad <= 28 Then
            intDebugLoad += 1
            btnDebugLoad.Text = ("Load question " + intDebugLoad.ToString)
        End If
    End Sub

    Private Sub btnDebugLoad_Click(sender As Object, e As EventArgs) Handles btnDebugLoad.Click
        Call CurrentQuestion(intDebugLoad)
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        btnStart.Visible = False
        grpbGame.Visible = True
        grpbGame.Enabled = True
        'Debug Mode
        If Control.ModifierKeys = Keys.Shift Then
            MessageBox.Show("Shift key was pressed along with click, Enabling debug mode.")
            grpbDebugMenu.Visible = True
        End If

        Call Game()

    End Sub

    Private Sub Game()

        'Random
        Dim random1 As New Random

        'Level state 
        Dim intCurrentLevel As Integer = 0
        Dim intCurrentQuestion As Integer = 0

        intMoneyHeld = 0

        btn5050.Enabled = True
        btnPoll.Enabled = True
        btnHint.Enabled = True

        Do Until intCurrentLevel = 16
            'Increment game level
            intCurrentLevel += 1

            blnWaitingForInput = False

            Select Case intCurrentLevel
                Case 1
                    intMoneyHeld = 0
                    intQuestionDifficulty = 1
                Case 2
                    intMoneyHeld = 200
                    lblPrize0.BackColor = Color.White
                    lblPrize200.BackColor = Color.LimeGreen
                Case 3
                    intMoneyHeld = 300
                    lblPrize200.BackColor = Color.White
                    lblPrize300.BackColor = Color.LimeGreen
                Case 4
                    intMoneyHeld = 500
                    intQuestionDifficulty = 2
                    lblPrize300.BackColor = Color.White
                    lblPrize500.BackColor = Color.LimeGreen
                Case 5
                    intMoneyHeld = 1000
                    lblPrize500.BackColor = Color.White
                    lblPrize1000.BackColor = Color.LimeGreen
                Case 6
                    intMoneyHeld = 2000
                    lblPrize1000.BackColor = Color.Orange
                    lblPrize2000.BackColor = Color.LimeGreen
                Case 7
                    intMoneyHeld = 4000
                    intQuestionDifficulty = 3
                    lblPrize2000.BackColor = Color.White
                    lblPrize4000.BackColor = Color.LimeGreen
                Case 8
                    intMoneyHeld = 8000
                    lblPrize4000.BackColor = Color.White
                    lblPrize8000.BackColor = Color.LimeGreen
                Case 9
                    intMoneyHeld = 16000
                    lblPrize8000.BackColor = Color.White
                    lblPrize16000.BackColor = Color.LimeGreen
                Case 10
                    intMoneyHeld = 32000
                    intQuestionDifficulty = 4
                    lblPrize16000.BackColor = Color.White
                    lblPrize32000.BackColor = Color.LimeGreen
                Case 11
                    intMoneyHeld = 64000
                    lblPrize32000.BackColor = Color.Orange
                    lblPrize64000.BackColor = Color.LimeGreen
                Case 12
                    intMoneyHeld = 125000
                    lblPrize64000.BackColor = Color.White
                    lblPrize125000.BackColor = Color.LimeGreen
                Case 13
                    intMoneyHeld = 250000
                    intQuestionDifficulty = 5
                    lblPrize125000.BackColor = Color.White
                    lblPrize250000.BackColor = Color.LimeGreen
                Case 14
                    intMoneyHeld = 500000
                    lblPrize250000.BackColor = Color.White
                    lblPrize500000.BackColor = Color.LimeGreen
                Case 15
                    intMoneyHeld = 1000000
                    lblPrize500000.BackColor = Color.White
                    lblPrize1000000.BackColor = Color.LimeGreen
            End Select

            lblMoneyHeld.Text = ("Current prize: $" + intMoneyHeld.ToString)

            'check if game is won
            If intMoneyHeld = 1000000 Then
                MessageBox.Show("Congratulations! You won and are now a millionaire!")
                Application.Restart()
            End If

            'Loads a question for the current difficulty tier
            Select Case intQuestionDifficulty
                Case 1
                    Call CurrentQuestion(random1.Next(1, 10))
                Case 2
                    Call CurrentQuestion(random1.Next(10, 17))
                Case 3
                    Call CurrentQuestion(random1.Next(17, 22))
                Case 4
                    Call CurrentQuestion(random1.Next(22, 27))
                Case 5
                    Call CurrentQuestion(random1.Next(27, 30))
            End Select



            ' Get the current question and its answers

            Dim CQCorrectAnswer As String = debugAnswer.Text
            Dim CQWrongAnswer1 As String = debugWrong1.Text
            Dim CQWrongAnswer2 As String = debugWrong2.Text
            Dim CQWrongAnswer3 As String = debugWrong3.Text

            ' Store answers in a list
            Dim answers As New List(Of String) From {CQCorrectAnswer, CQWrongAnswer1, CQWrongAnswer2, CQWrongAnswer3}
            Shuffle(answers)
            lblAnswer1Text.Text = answers(0)
            lblAnswer2Text.Text = answers(1)
            lblAnswer3Text.Text = answers(2)
            lblAnswer4Text.Text = answers(3)

            If answers(0) = CQCorrectAnswer Then
                intCorrectAnswer = 1
            ElseIf answers(1) = CQCorrectAnswer Then
                intCorrectAnswer = 2
            ElseIf answers(2) = CQCorrectAnswer Then
                intCorrectAnswer = 3
            ElseIf answers(3) = CQCorrectAnswer Then
                intCorrectAnswer = 4
            End If

            blnWaitingForInput = True

            btnConfirm.Enabled = False

            btnAnswer1.Enabled = True
            btnAnswer2.Enabled = True
            btnAnswer3.Enabled = True
            btnAnswer4.Enabled = True

            AddHandler btnAnswer1.Click, AddressOf AnswerSelected
            AddHandler btnAnswer2.Click, AddressOf AnswerSelected
            AddHandler btnAnswer3.Click, AddressOf AnswerSelected
            AddHandler btnAnswer4.Click, AddressOf AnswerSelected

            AddHandler btn5050.Click, AddressOf call5050
            AddHandler btnPoll.Click, AddressOf callPoll
            AddHandler btnHint.Click, AddressOf callHint


            While blnWaitingForInput = True
                Application.DoEvents() 'This allows the UI to remain responsive while waiting
            End While

            RemoveHandler btnAnswer1.Click, AddressOf AnswerSelected
            RemoveHandler btnAnswer2.Click, AddressOf AnswerSelected
            RemoveHandler btnAnswer3.Click, AddressOf AnswerSelected
            RemoveHandler btnAnswer4.Click, AddressOf AnswerSelected

            RemoveHandler btn5050.Click, AddressOf call5050
            RemoveHandler btnPoll.Click, AddressOf callPoll
            RemoveHandler btnHint.Click, AddressOf callHint

        Loop ' End of game loop
    End Sub

    Private Sub btnAnswer_Click(sender As Object, e As EventArgs) Handles btnAnswer1.Click, btnAnswer2.Click, btnAnswer3.Click, btnAnswer4.Click
        ' Determine which button was clicked
        Dim selectedAnswerbtn As Integer = Integer.Parse(sender.Name.Replace("btnAnswer", ""))

        ' Check if the selected answer is correct
        If selectedAnswerbtn = intCorrectAnswer Then
            blnCorrectAnswerSelected = True
        Else
            blnCorrectAnswerSelected = False
        End If
    End Sub

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        ' Check if the selected answer is correct
        If blnCorrectAnswerSelected = True Then
            MessageBox.Show("Correct answer!")
            blnWaitingForInput = False
        Else
            MessageBox.Show("Wrong answer! You walk away with $" + intMoneyHeld.ToString + ". Thank you for playing.")
            Application.Restart()
        End If

    End Sub

    Private Sub Shuffle(Of T)(list As IList(Of T))
        Dim random2 As New Random
        Dim n As Integer = list.Count
        While n > 1
            n -= 1
            Dim k As Integer = random2.Next(n + 1)
            Dim value As T = list(k)
            list(k) = list(n)
            list(n) = value
        End While
    End Sub

    Private Sub AnswerSelected(sender As Object, e As EventArgs)
        'Disable the answer buttons after selection
        btnAnswer1.Enabled = False
        btnAnswer2.Enabled = False
        btnAnswer3.Enabled = False
        btnAnswer4.Enabled = False

        'Enable the confirm button
        btnConfirm.Enabled = True

        'Store the selected answer
        If sender Is btnAnswer1 Then
            intAnswerSelected = 1
        ElseIf sender Is btnAnswer2 Then
            intAnswerSelected = 2
        ElseIf sender Is btnAnswer3 Then
            intAnswerSelected = 3
        ElseIf sender Is btnAnswer4 Then
            intAnswerSelected = 4
        End If
    End Sub

    Private Sub call5050(sender As Object, e As EventArgs)
        Dim str50501 As String = debug5050.Text
        Dim str50502 As String = debugAnswer.Text
        Dim random3 As New Random
        Dim order As Integer = random3.Next(6)
        If btn5050.Enabled = True Then
            If order = order Mod 2 Then
                MessageBox.Show("The answer is either " + str50501 + " or " + str50502 + ".")
            Else
                MessageBox.Show("The answer is either " + str50502 + " or " + str50501 + ".")
            End If
        End If
        btn5050.Enabled = False
    End Sub

    Private Sub callPoll(sender As Object, e As EventArgs)
        If btnPoll.Enabled = True Then
            MessageBox.Show("This the the percentage of the audience who each voted for a,b,c,d: " + debugPoll.Text)
        End If
        btnPoll.Enabled = False
    End Sub

    Private Sub callHint(sender As Object, e As EventArgs)
        If btnHint.Enabled = True Then
            MessageBox.Show(debugHint.Text)
        End If
        btnHint.Enabled = False
    End Sub

    Private Sub btnTakeMoney_Click(sender As Object, e As EventArgs) Handles btnTakeMoney.Click
        MessageBox.Show("You walk away with $" + intMoneyHeld.ToString + ". Thank you for playing.")
        Application.Restart()
    End Sub
End Class
