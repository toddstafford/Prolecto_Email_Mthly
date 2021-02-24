
Imports System.Text
Imports System.Data.SqlClient
Imports System.Net.Mail


Module Program
    Private dcFINKSRPT As String = "Data Source=RAYSQL;Initial Catalog=FINKS_RPT;Persist Security Info=True;User ID=sa;Password=314159"
    Sub Main()

        Dim dt As New DateTime
        Dim dow As Int16

        dt = Now()

        dow = CInt(dt.DayOfWeek)

        If dow = 0 Then
            DoItemHistory()
        ElseIf dow = 6 Then
            DoHierarchyHistory()
        End If


        DoStoreHistory()
        DoVendorHistory()
        DoTeamHistory()


    End Sub

    Sub DoHierarchyHistory()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_HIERARCH_HISTORY_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure


        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder

        FILE_NAME = "\\garnet\ProlectoSends\FINKS_HIERARCHY_HISTORY" & ".csv"

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
            Dim thereport As New Attachment(FILE_NAME.ToString)
            SendMail(thereport, (Now()), "HIERARCHY_HISTORY")
        End Try
    End Sub
    Sub DoStoreHistory()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_STORE_HISTORY_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure

        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder


        FILE_NAME = "\\garnet\ProlectoSends\FINKS_STORE_HISTORY" & ".csv"
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
            Dim thereport As New Attachment(FILE_NAME.ToString)
            SendMail(thereport, (Now()), "STORE_HISTORY")

        End Try
    End Sub
    Sub DoVendorHistory()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_VENDOR_HISTORY_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure

        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder


        FILE_NAME = "\\garnet\ProlectoSends\FINKS_VENDOR_HISTORY" & ".csv"
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
            Dim thereport As New Attachment(FILE_NAME.ToString)
            SendMail(thereport, (Now()), "VENDOR_HISTORY")

        End Try
    End Sub
    Sub DoItemHistory()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_ITEM_HISTORY_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure

        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder


        FILE_NAME = "\\garnet\ProlectoSends\FINKS_ITEM_HISTORY" & ".csv"
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
            Dim thereport As New Attachment(FILE_NAME.ToString)
            SendMail(thereport, (Now()), "Item History")

        End Try
    End Sub
    Sub DoTeamHistory()
        Dim Conn As SqlConnection = New SqlConnection(dcFINKSRPT)
        Dim CMD As SqlCommand = New SqlCommand("Prolecto_TEAM_HISTORY_Feed", Conn)
        CMD.CommandType = Data.CommandType.StoredProcedure

        Conn.Open()

        Dim myReader As SqlDataReader = CMD.ExecuteReader()
        Dim dt As New Data.DataTable
        dt.Load(myReader)


        myReader.Close()

        Dim FILE_NAME As String
        Dim CsvFile As New StringBuilder


        FILE_NAME = "\\garnet\ProlectoSends\FINKS_TEAM_HISTORY" & ".csv"
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
            Dim thereport As New Attachment(FILE_NAME.ToString)
            SendMail(thereport, (Now()), "Team History")

        End Try
    End Sub
    Private Sub SendMail(ByVal fname As Attachment, rundate As Date, subj As String)

        Try
            Dim myfrom As New MailAddress("tstafford@finks.com") 'emails.5404883.873.81ce8c857b@5404883.email.netsuite.com
            Dim myto As New MailAddress("emails.5404883.873.81ce8c857b@5404883.email.netsuite.com")
            'Dim myto As New MailAddress("stffrdtodd@gmail.com")
            'Dim cc2 As New MailAddress("salesadmin@rusa.com")
            Dim cc2 As New MailAddress("tstafford@finks.com")



            Dim rolexMail As New MailMessage(myfrom, myto)
            With rolexMail
                .Subject = "Prolecto -" & rundate & " " & subj
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
