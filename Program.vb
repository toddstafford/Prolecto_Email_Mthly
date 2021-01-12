
Imports System.Text
Imports System.Data.SqlClient
Imports System.Net.Mail


Module Program
    Private dcFINKSRPT As String = "Data Source=RAYSQL;Initial Catalog=FINKS_RPT;Persist Security Info=True;User ID=sa;Password=314159"
    Sub Main()
        DoDaily()
        DoTeam()
        DoCPH()
    End Sub

    Sub DoDaily()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_DailySalesDetail_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure


        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder


        FILE_NAME = "\\garnet\ProlectoSends\FINKS_DAILY_SALES_DETAIL" & ".csv"

        'FILE_NAME = "C:\Users\Tstafford\Documents\YURMAN_FINKS_" & SaturdayDate.ToString("yyyyMMdd") & ".csv"
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
        Try

            ' first write a line with the columns name

            Dim sep As String = ""

            Dim builder As New System.Text.StringBuilder

            For Each col As Data.DataColumn In dt.Columns

                builder.Append(sep).Append(col.ColumnName)

                sep = ","
            Next

            objWriter.WriteLine(builder.ToString())

            ' then write all the rows

            For Each row As Data.DataRow In dt.Rows

                sep = ""

                builder = New System.Text.StringBuilder

                For Each col As Data.DataColumn In dt.Columns

                    builder.Append(sep).Append(row(col.ColumnName))

                    sep = ","

                Next

                objWriter.WriteLine(builder.ToString())

            Next

        Finally

            If Not objWriter Is Nothing Then objWriter.Close()
            Dim DAILYreport As New Attachment(FILE_NAME.ToString)
            SendMail(DAILYreport, (Now()))
        End Try
    End Sub
    Sub DoTeam()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_DailyTeamSales_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure

        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder


        FILE_NAME = "\\garnet\ProlectoSends\FINKS_DAILY_TEAM_SALES" & ".csv"
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
        Try

            ' first write a line with the columns name

            Dim sep As String = ""

            Dim builder As New System.Text.StringBuilder

            For Each col As Data.DataColumn In dt.Columns

                builder.Append(sep).Append(col.ColumnName)

                sep = ","
            Next

            objWriter.WriteLine(builder.ToString())

            ' then write all the rows

            For Each row As Data.DataRow In dt.Rows

                sep = ""

                builder = New System.Text.StringBuilder

                For Each col As Data.DataColumn In dt.Columns

                    builder.Append(sep).Append(row(col.ColumnName))

                    sep = ","

                Next

                objWriter.WriteLine(builder.ToString())

            Next

        Finally

            If Not objWriter Is Nothing Then objWriter.Close()
            Dim TEAMreport As New Attachment(FILE_NAME.ToString)
            SendMail(TEAMreport, (Now()))

        End Try
    End Sub
    Sub DoCPH()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_CPH_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure

        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder


        FILE_NAME = "\\garnet\ProlectoSends\FINKS_CPH" & ".csv"
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
        Try

            ' first write a line with the columns name

            Dim sep As String = ""

            Dim builder As New System.Text.StringBuilder

            For Each col As Data.DataColumn In dt.Columns

                builder.Append(sep).Append(col.ColumnName)

                sep = ","
            Next

            objWriter.WriteLine(builder.ToString())

            ' then write all the rows

            For Each row As Data.DataRow In dt.Rows

                sep = ""

                builder = New System.Text.StringBuilder

                For Each col As Data.DataColumn In dt.Columns

                    builder.Append(sep).Append(row(col.ColumnName))

                    sep = ","

                Next

                objWriter.WriteLine(builder.ToString())

            Next

        Finally

            If Not objWriter Is Nothing Then objWriter.Close()
            Dim CPHreport As New Attachment(FILE_NAME.ToString)
            SendMail(CPHreport, (Now()))

        End Try
    End Sub

    Private Sub SendMail(ByVal fname As Attachment, rundate As Date)

        Try
            Dim myfrom As New MailAddress("tstafford@finks.com") 'emails.5404883.873.81ce8c857b@5404883.email.netsuite.com
            Dim myto As New MailAddress("emails.5404883.873.81ce8c857b@5404883.email.netsuite.com")
            'Dim myto As New MailAddress("stffrdtodd@gmail.com")
            'Dim cc2 As New MailAddress("salesadmin@rusa.com")
            Dim cc2 As New MailAddress("tstafford@finks.com")

            Dim rolexMail As New MailMessage(myfrom, myto)
            With rolexMail
                .Subject = "Prolecto -" & rundate
                .Body = "Fink's Jewelers"
                .Attachments.Add(fname)
                .To.Add(cc2)
                '.To.Add(cc3)
            End With

            Dim mySmtp As New SmtpClient
            mySmtp.Host = "192.168.0.194"
            mySmtp.Send(rolexMail)
        Catch ex As Exception
            Console.WriteLine(ex.Message)

        End Try
    End Sub
End Module
