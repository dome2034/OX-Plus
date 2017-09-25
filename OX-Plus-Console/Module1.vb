Module Module1
    Dim Board As String
    Dim Mak(2, 2), IsContinue, First As Char
    Dim Score(2, 2) As Integer
    Dim Player, HaveWinner, Round, inputA, inputB, Point As UInteger
    Dim Playing As Boolean

    Sub Main()
        Playing = True
        Player = 1 '1 = O turn, 2 = X turn
        While Playing = True
            initBoard()
            Console.Write(Board)
            While HaveWinner = 0 And Round < 9
                Try
                    If Player = 1 Then
                        Console.Write("o (A) : ")
                        inputA = Console.ReadLine() - 1
                        Console.Write("o (B) : ")
                        inputB = Console.ReadLine() - 1
                    Else
                        Console.Write("x (A) : ")
                        inputA = Console.ReadLine() - 1
                        Console.Write("x (B) : ")
                        inputB = Console.ReadLine() - 1
                    End If

                    If Mak(inputA, inputB) = " " Then
                        If Player = 1 Then
                            Mak(inputA, inputB) = "O"
                            Player = 2
                        Else
                            Mak(inputA, inputB) = "X"
                            Player = 1
                        End If
                        Round = Round + 1
                    Else
                        Console.WriteLine("Invalid move, please choose again!")
                        Console.ReadLine()
                    End If
                Catch ex As Exception
                    Console.WriteLine("Invalid move, please choose again!")
                    Console.ReadLine()
                End Try
                '-------- end input state -------------------------
                '--------------------------------------------------
                Console.Clear()
                drawBoard(Board, Mak)
                HaveWinner = checkWin(Mak)
            End While

            If HaveWinner = 1 Then
                Console.Write("o won with {0} Points", Point)
                Console.ReadLine()
            ElseIf HaveWinner = 2 Then
                Console.Write("x won with {0} Points", Point)
                Console.ReadLine()
            Else
                Console.Write("DRAW !!!")
                Console.ReadLine()
            End If

            checkTop5(Point)
            showTop5Records()
            showRecordsAll()
            Console.WriteLine("-------------------------------------------")
            While IsContinue <> "y" And IsContinue <> "n"
                Console.Write("Do you want to play again (y/n) ? ")
                IsContinue = Console.ReadLine().ToLower()
            End While
            If IsContinue = "y" Then
                Playing = True
            Else
                Playing = False
                Exit While
            End If

            While First <> "x" And First <> "o"
                Console.Write("Who's first (o/x) ? ")
                First = Console.ReadLine().ToLower()
            End While
            If First = "o" Then
                Player = 1
            Else
                Player = 2
            End If
            IsContinue = Nothing
            First = Nothing
            Console.Clear()

        End While
        Console.Clear()
        Console.WriteLine("-------------------------------------------")
        Console.WriteLine("---------------- Good Bye -----------------")
        Console.WriteLine("-------------------------------------------")
        Console.ReadLine()

    End Sub

    Sub initBoard()

        Board = "   B1  B2  B3 " + vbCrLf
        Board += "A1   |   |   " + vbCrLf
        Board += "  ---+---+---" + vbCrLf
        Board += "A2   |   |   " + vbCrLf
        Board += "  ---+---+---" + vbCrLf
        Board += "A3   |   |   " + vbCrLf

        For i = 0 To 2
            For j = 0 To 2
                Mak(i, j) = " "
            Next
        Next

        For i = 0 To 2
            For j = 0 To 2
                Score(i, j) = CInt(8 * Rnd()) + 1
                Randomize()
            Next
        Next

        Round = 0
        HaveWinner = 0
        Point = 0
    End Sub
    Sub drawBoard(Board As String, Mak As Char(,))
        Board = "   B1  B2  B3 " + vbCrLf
        Board += "A1 {0} | {1} | {2} " + vbCrLf
        Board += "  ---+---+---" + vbCrLf
        Board += "A2 {3} | {4} | {5} " + vbCrLf
        Board += "  ---+---+---" + vbCrLf
        Board += "A3 {6} | {7} | {8} " + vbCrLf
        'Console.Write(Board, Score(0, 0), Score(0, 1), Score(0, 2), Score(1, 0), Score(1, 1), Score(1, 2), Score(2, 0), Score(2, 1), Score(2, 2))
        'Console.WriteLine("-----------------------")

        Console.Write(Board, Mak(0, 0), Mak(0, 1), Mak(0, 2), Mak(1, 0), Mak(1, 1), Mak(1, 2), Mak(2, 0), Mak(2, 1), Mak(2, 2))
        Console.WriteLine("-----------------------")
    End Sub
    Function checkWin(Mak As Char(,)) As Integer

        If Mak(0, 0) + Mak(0, 1) + Mak(0, 2) = "OOO" Then '1
            Point = Score(0, 0) + Score(0, 1) + Score(0, 2)
            Return 1
        ElseIf Mak(1, 0) + Mak(1, 1) + Mak(1, 2) = "OOO" Then '2
            Point = Score(1, 0) + Score(1, 1) + Score(1, 2)
            Return 1
        ElseIf Mak(2, 0) + Mak(2, 1) + Mak(2, 2) = "OOO" Then '3
            Point = Score(2, 0) + Score(2, 1) + Score(2, 2)
            Return 1
        ElseIf Mak(0, 0) + Mak(1, 0) + Mak(2, 0) = "OOO" Then '4
            Point = Score(0, 0) + Score(1, 0) + Score(2, 0)
            Return 1
        ElseIf Mak(0, 1) + Mak(1, 1) + Mak(2, 1) = "OOO" Then '5
            Point = Score(0, 1) + Score(1, 1) + Score(2, 1)
            Return 1
        ElseIf Mak(0, 2) + Mak(1, 2) + Mak(2, 2) = "OOO" Then '6
            Point = Score(0, 2) + Score(1, 2) + Score(2, 2)
            Return 1
        ElseIf Mak(0, 0) + Mak(1, 1) + Mak(2, 2) = "OOO" Then '7
            Point = Score(0, 0) + Score(1, 1) + Score(2, 2)
            Return 1
        ElseIf Mak(0, 2) + Mak(1, 1) + Mak(2, 0) = "OOO" Then '8
            Point = Score(0, 2) + Score(1, 1) + Score(2, 0)
            Return 1
        ElseIf Mak(0, 0) + Mak(0, 1) + Mak(0, 2) = "XXX" Then '1
            Point = Score(0, 0) + Score(0, 1) + Score(0, 2)
            Return 2
        ElseIf Mak(1, 0) + Mak(1, 1) + Mak(1, 2) = "XXX" Then '2
            Point = Score(1, 0) + Score(1, 1) + Score(1, 2)
            Return 2
        ElseIf Mak(2, 0) + Mak(2, 1) + Mak(2, 2) = "XXX" Then '3
            Point = Score(2, 0) + Score(2, 1) + Score(2, 2)
            Return 2
        ElseIf Mak(0, 0) + Mak(1, 0) + Mak(2, 0) = "XXX" Then '4
            Point = Score(0, 0) + Score(1, 0) + Score(2, 0)
            Return 2
        ElseIf Mak(0, 1) + Mak(1, 1) + Mak(2, 1) = "XXX" Then '5
            Point = Score(0, 1) + Score(1, 1) + Score(2, 1)
            Return 2
        ElseIf Mak(0, 2) + Mak(1, 2) + Mak(2, 2) = "XXX" Then '6
            Point = Score(0, 2) + Score(1, 2) + Score(2, 2)
            Return 2
        ElseIf Mak(0, 0) + Mak(1, 1) + Mak(2, 2) = "XXX" Then '7
            Point = Score(0, 0) + Score(1, 1) + Score(2, 2)
            Return 2
        ElseIf Mak(0, 2) + Mak(1, 1) + Mak(2, 0) = "XXX" Then '8
            Point = Score(0, 2) + Score(1, 1) + Score(2, 0)
            Return 2
        Else
            Return 0
        End If

    End Function
    Sub checkTop5(Score As Integer)
        Dim sqCon As New SqlClient.SqlConnection("Server=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\OX_DB.mdf;Database=OX_DB;Trusted_Connection=Yes;")
        Dim cmdSelect, cmdInsert As SqlClient.SqlCommand
        Dim SQL_Query, NameTop5 As String
        Dim CheckTable As DataTable = New DataTable()
        Dim HaveNewRecord As Boolean

        SQL_Query = "Select TOP 5 name,score From dbo.Score ORDER BY score DESC"
        cmdSelect = New SqlClient.SqlCommand(SQL_Query, sqCon)

        sqCon.Open()
        Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(cmdSelect)
        da.Fill(CheckTable)
        sqCon.Close()

        HaveNewRecord = False
        For Each ScoreRow As DataRow In CheckTable.Rows
            If Score > ScoreRow.Item("score") Then
                HaveNewRecord = True
                Exit For
            End If
        Next

        If CheckTable.Rows.Count < 5 Or HaveNewRecord Then
            Console.WriteLine("Congratulation !! you are top 5.")
            Console.Write("please input your name : ")
            NameTop5 = Console.ReadLine()
            SQL_Query = "INSERT INTO dbo.Score VALUES ('" & NameTop5 & "'," & Score.ToString() & ")"
            sqCon.Open()
            cmdInsert = New SqlClient.SqlCommand(SQL_Query, sqCon)
            cmdInsert.ExecuteNonQuery()
            sqCon.Close()
        End If
    End Sub

    Sub showTop5Records()
        Dim sqCon As New SqlClient.SqlConnection("Server=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\OX_DB.mdf;Database=OX_DB;Trusted_Connection=Yes;")
        Dim cmdSelect As SqlClient.SqlCommand
        Dim SQL_Query As String
        Dim CheckTable As DataTable = New DataTable()
        Dim Order As Integer
        Order = 1

        SQL_Query = "Select TOP 5 name,score From dbo.Score ORDER BY score DESC "
        cmdSelect = New SqlClient.SqlCommand(SQL_Query, sqCon)

        sqCon.Open()
        Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(cmdSelect)
        da.Fill(CheckTable)
        sqCon.Close()

        For Each ScoreRow As DataRow In CheckTable.Rows
            Console.WriteLine("{0}. Name : {1} Have Score : {2} points", Order, ScoreRow.Item("name"), ScoreRow.Item("score"))
            Order = Order + 1
        Next
        Console.WriteLine("-------------------------------------------")
    End Sub
    Sub showRecordsAll()
        Dim sqCon As New SqlClient.SqlConnection("Server=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\OX_DB.mdf;Database=OX_DB;Trusted_Connection=Yes;")
        Dim cmdSelect As SqlClient.SqlCommand
        Dim SQL_Query As String
        Dim CheckTable As DataTable = New DataTable()
        Dim Order As Integer
        Order = 1

        SQL_Query = "Select name,score From dbo.Score "
        cmdSelect = New SqlClient.SqlCommand(SQL_Query, sqCon)

        sqCon.Open()
        Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(cmdSelect)
        da.Fill(CheckTable)
        sqCon.Close()

        For Each ScoreRow As DataRow In CheckTable.Rows
            Console.WriteLine("{0}. Name : {1} Have Score : {2} points", Order, ScoreRow.Item("name"), ScoreRow.Item("score"))
            Order = Order + 1
        Next
    End Sub
End Module
