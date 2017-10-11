'********************************************************************************************
'Views required in SQL:
'   vMR_ItemPrime
'   vMR_ItemToReplen
'   vMR_AllBulksToReplen
'   vMR_OldestBulktoReplen
'   vMR_ItemToPutaway

'Config File
'-----------
'need a settings file : "C:\dash\ReplenSettings.txt"
'in this format:
'SQLServer{\instance}
'DB Name
'User 
'Password
'Sample
'------
'ianmac-lenovo\ianmac
'Dexterity
'sa
'Chatburn441977

' Directory Needed on Server:
'C:\Dash\FromWMS\ReplenBulk 
'C:\Dash\FromWMS\ReplenPutaway 
' and z: mapping to Dexterity server c:\Dash\FromWMS which also requires sharing

'NiceLabel
'Ensure triggers delete on print as file name is static and will overright each time prints are requested


Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient

Public Class frmManReplen
    'pubs:
    Public txtDBName As String
    Public txtServer As String
    Public txtUser As String
    Public txtPassword As String
    Public intLineCount As Int64
    Public intDetailLineCount As Integer


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        End
    End Sub

    Private Sub frmManReplen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Get DB Settings
        Call getDBSettings()

        'Clear ListView
        lvBulk.Items.Clear()
        lvPutaway.Items.Clear()

        'Load List Views

        'Bulkpicks:
        Call GetBulkPicksLines()

        'Putaways:
        Call GetPutawayLines()



    End Sub
    Sub GetBulkPicksLines()


        'SQL connection string
        Dim connStringCAL As String = _
           "server = " & txtServer & ";" _
         & "user id = " & txtUser & ";" _
         & "password = " & txtPassword & ";" _
         & "database = " & txtDBName

        Dim sqlQuery As String
        'query to get header
        sqlQuery = _
        "SELECT   " & _
         "      locationcode as Location,   " & _
         "      itemsdesc as Item,   " & _
         "      OnHandQuantity as BulkQty, " & _
         "      PrimeLocation _ " & _
         "FROM   " & _
         "     vMR_OldestBulkToReplen " & _
         "ORDER By Picksequence"

        'Console.Write(sqlDetailQuery)
        Dim txtLocation As String = ""
        Dim txtItem As String = ""
        Dim txtBulkQty As String = ""
        Dim txtPrimeLocation As String = ""

        'create connection
        Dim connCAL As SqlConnection = New SqlConnection(connStringCAL)

        'create commands
        Dim cmdQryCAL As SqlCommand = New SqlCommand(sqlQuery, connCAL)

        'WMS Data
        Try
            'open connection
            connCAL.Open()

            'Execute header query
            Dim rdrCAL As SqlDataReader = cmdQryCAL.ExecuteReader()

            While rdrCAL.Read

                'read line
                txtLocation = (rdrCAL.GetSqlString(0).ToString)
                txtItem = (rdrCAL.GetSqlString(1).ToString)
                txtBulkQty = (rdrCAL.GetSqlDecimal(2).ToString)
                txtPrimeLocation = (rdrCAL.GetSqlString(3).ToString)

                'push into listview
                lvBulk.Items.Add(txtLocation.ToString)
                lvBulk.Items(lvBulk.Items.Count - 1).SubItems.Add(txtItem.ToString)
                lvBulk.Items(lvBulk.Items.Count - 1).SubItems.Add(txtBulkQty.ToString)
                lvBulk.Items(lvBulk.Items.Count - 1).SubItems.Add(txtPrimeLocation.ToString)


            End While

        Catch ex As Exception
            'display error
            Console.WriteLine("Error SQL  : " & ex.ToString)

        Finally
            'close connection
            connCAL.Close()
        End Try
    End Sub
    Sub GetPutawayLines()


        'SQL connection string
        Dim connStringCAL As String = _
           "server = " & txtServer & ";" _
         & "user id = " & txtUser & ";" _
         & "password = " & txtPassword & ";" _
         & "database = " & txtDBName

        Dim sqlQuery As String
        'query to get header
        sqlQuery = _
        "SELECT	Item, " & _
         "	PrimaryLocation, " & _
         "	Picksequence  " & _
         "FROM     vmr_itemtoputaway " & _
         "ORDER BY Picksequence"

        'Console.Write(sqlDetailQuery)
        Dim txtPriLocation As String = ""
        Dim txtItem As String = ""
        Dim txtPickSeq As String = ""

        'create connection
        Dim connCAL As SqlConnection = New SqlConnection(connStringCAL)

        'create commands
        Dim cmdQryCAL As SqlCommand = New SqlCommand(sqlQuery, connCAL)

        'WMS Data
        Try
            'open connection
            connCAL.Open()

            'Execute header query
            Dim rdrCAL As SqlDataReader = cmdQryCAL.ExecuteReader()

            While rdrCAL.Read

                'read line
                txtItem = (rdrCAL.GetSqlString(0).ToString)
                txtPriLocation = (rdrCAL.GetSqlString(1).ToString)
                txtPickSeq = (rdrCAL.GetSqlInt32(2).ToString)

                'push into listview
                lvPutaway.Items.Add(txtItem.ToString)
                lvPutaway.Items(lvPutaway.Items.Count - 1).SubItems.Add(txtPriLocation.ToString)
                lvPutaway.Items(lvPutaway.Items.Count - 1).SubItems.Add(txtPickSeq.ToString)

            End While

        Catch ex As Exception
            'display error
            Console.WriteLine("Error SQL  : " & ex.ToString)

        Finally
            'close connection
            connCAL.Close()
        End Try
    End Sub

    Sub getDBSettings()

        Dim FILE_NAME As String = "C:\dash\ReplenSettings.txt"


        If System.IO.File.Exists(FILE_NAME) = True Then
            Dim objReader As New System.IO.StreamReader(FILE_NAME)
            txtServer = objReader.ReadLine() '& vbNewLine
            txtDBName = objReader.ReadLine() '& vbNewLine
            txtUser = objReader.ReadLine() '& vbNewLine
            txtPassword = objReader.ReadLine() '& vbNewLine
            objReader.Close()

        Else
            Console.Write("Error Openeing Settings File - Please check c:\Dash\ReplenSettings.txt is there")
        End If
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        'Get DB Settings
        Call getDBSettings()

        'Clear ListView
        lvBulk.Items.Clear()
        lvPutaway.Items.Clear()

        'Load List Views

        'Bulkpicks:
        Call GetBulkPicksLines()

        'Putaways:
        Call GetPutawayLines()


    End Sub

    Private Sub btnPrintBulk_Click(sender As Object, e As EventArgs) Handles btnPrintBulk.Click
        'Loop through contents of lvbulkpicks and spit it out
        Dim strLocation As String = ""
        Dim strItem As String = ""
        Dim strBulkQty As String = ""
        Dim strPrimeLocation As String = ""
        Dim lineout As String = ""
        Dim pagecount As Integer = 1
        Dim linecount As Integer = 0
        Dim sSource As String
        Dim sTarget As String
        sSource = "c:\dash\tempBulkReplen"
        sTarget = "z:\ReplenBulk\BulkReplen"

        'create empty temp file
        File.Create("C:\dash\tempBulkReplen" & pagecount.ToString & ".txt").Dispose()

        'prepare temp file
        Dim tmpBulkReplen As String = "C:\dash\tempBulkReplen" & pagecount.ToString & ".txt"

        Dim objWriter As New System.IO.StreamWriter(tmpBulkReplen)
        For Each item As ListViewItem In Me.lvBulk.Items
            linecount = linecount + 1
            strLocation = item.Text
            strItem = item.SubItems.Item(1).Text
            strBulkQty = item.SubItems.Item(2).Text
            strPrimeLocation = item.SubItems.Item(3).Text
            If (linecount Mod 20) = 0 Then
                'close file - reached line limit
                objWriter.Close()

                'copy completed file to print directory
                sSource = "c:\dash\tempBulkReplen" & pagecount.ToString & ".txt"
                sTarget = "z:\ReplenBulk\BulkReplen" & pagecount.ToString & ".txt"
                File.Copy(sSource, sTarget, True)
                pagecount = pagecount + 1

                'create empty temp file
                File.Create("C:\dash\tempBulkReplen" & pagecount.ToString & ".txt").Dispose()

                'prepare temp file
                tmpBulkReplen = "C:\dash\tempBulkReplen" & pagecount.ToString & ".txt"
                objWriter = New System.IO.StreamWriter(tmpBulkReplen)
            End If
            lineout = strLocation & "|" & strItem & "|" & strBulkQty & "|" & strPrimeLocation & "|" & pagecount.ToString & vbCrLf
            objWriter.Write(lineout)

        Next
        objWriter.Close()
        'reset filenames
        sSource = "c:\dash\tempBulkReplen"
        sTarget = "z:\ReplenBulk\BulkReplen"
        'make sure line count not end of page and already been copied over to
        If (linecount Mod 20) <> 0 Then
            'copy completed file to print directory
            sSource = "c:\dash\tempBulkReplen" & pagecount.ToString & ".txt"
            sTarget = "z:\ReplenBulk\BulkReplen" & pagecount.ToString & ".txt"
            File.Copy(sSource, sTarget, True)
        End If

    End Sub

    Private Sub btnPutaway_Click(sender As Object, e As EventArgs) Handles btnPutaway.Click
        'Loop through contents of lvputaway and spit it out
        Dim strItem As String = ""
        Dim strPriLocation As String = ""
        Dim strSequence As String = ""
        Dim lineout As String = ""
        Dim pagecount As Integer = 1
        Dim linecount As Integer = 0
        Dim PrintedLastPage As Boolean = False
        Dim sSource As String
        Dim sTarget As String
        sSource = "c:\dash\tempPutawayReplen"
        sTarget = "z:\ReplenPutaway\PutawayReplen"

        'create empty temp file
        File.Create("C:\dash\tempPutawayReplen" & pagecount.ToString & ".txt").Dispose()

        'prepare temp file
        Dim tmpPutawayReplen As String = "C:\dash\tempPutawayReplen" & pagecount.ToString & ".txt"

        Dim objWriter As New System.IO.StreamWriter(tmpPutawayReplen)
        For Each item As ListViewItem In Me.lvPutaway.Items
            linecount = linecount + 1
            strItem = item.Text
            strPriLocation = item.SubItems.Item(1).Text
            strSequence = item.SubItems.Item(2).Text
            'reached a new page
            If (linecount Mod 20) = 0 Then
                'close file - reached line limit
                objWriter.Close()

                'copy completed file to print directory
                sSource = "c:\dash\tempPutawayReplen" & pagecount.ToString & ".txt"
                sTarget = "z:\ReplenPutaway\PutawayReplen" & pagecount.ToString & ".txt"
                File.Copy(sSource, sTarget, True)

                pagecount = pagecount + 1
                'create empty temp file
                File.Create("C:\dash\tempPutawayReplen" & pagecount.ToString & ".txt").Dispose()
                'prepare temp file
                tmpPutawayReplen = "C:\dash\tempPutawayReplen" & pagecount.ToString & ".txt"
                objWriter = New System.IO.StreamWriter(tmpPutawayReplen)
            End If
            lineout = strItem & "|" & strPriLocation & "|" & strSequence & "|" & pagecount.ToString & vbCrLf
            objWriter.Write(lineout)

        Next
        objWriter.Close()
        'reset filenames
        sSource = "c:\dash\tempPutawayReplen"
        sTarget = "z:\ReplenPutaway\PutawayReplen"
        'make sure line count not end of page and already been copied over to
        If (linecount Mod 20) <> 0 Then
            'copy completed file to print directory
            sSource = "c:\dash\tempPutawayReplen" & pagecount.ToString & ".txt"
            sTarget = "z:\ReplenPutaway\PutawayReplen" & pagecount.ToString & ".txt"
            File.Copy(sSource, sTarget, True)
        End If
    End Sub
End Class
