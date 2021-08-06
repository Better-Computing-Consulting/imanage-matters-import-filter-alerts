Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail
Module Module1
    Public intputfile As String = ""
    Public outputfile As String = ""
    Sub Main()
        If My.Application.CommandLineArgs.Count > 0 Then
            Dim downloadfile As String = My.Application.CommandLineArgs(0)
            Dim finfo As New FileInfo(downloadfile)
            If finfo.Exists Then
                Dim filteredfile As String = finfo.DirectoryName & "\Filtered." & finfo.Name
                'Console.WriteLine(filteredfile)
                FilterOutExistingMatters(downloadfile, filteredfile)
            Else
                Console.WriteLine(downloadfile & " does not exits at the specified location")
            End If
        Else
            Console.WriteLine("This application needs the path of the DMS import file to filter as a command line argument")
        End If
        Console.WriteLine("done")
    End Sub

    Sub FilterOutExistingMatters(downloadfilepath As String, filteredfilepath As String)
        Dim newmatters As New List(Of String)
        Dim existingmatters As New List(Of String)
        Dim downloadfile As String = downloadfilepath
        Dim filteredfile As String = filteredfilepath
        Dim connString1 As String = "Data Source=sql;Initial Catalog=iManage_Active;Integrated Security=SSPI"
        Dim queryString As String = "SELECT distinct c1alias, c2alias FROM MHGROUP.DOCMASTER WHERE C_ALIAS = 'WEBDOC' and C1ALIAS is not NULL order by c1alias"
        Using conn As New SqlConnection(connString1)
            Dim cmd As New SqlCommand(queryString, conn)
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader()
            If r.HasRows Then
                Try
                    While r.Read
                        Dim cnum As String = Trim(r("c1alias"))
                        Dim mnum As String = Trim(r("c2alias"))
                        Dim amatter As String = cnum.ToUpper & "," & mnum.ToUpper
                        existingmatters.Add(amatter)
                    End While
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            End If
        End Using
        Using sr As StreamReader = New StreamReader(downloadfile)
            Do While sr.Peek >= 0
                Try
                    Dim s As String = sr.ReadLine.ToUpper
                    If s.Length > 10 Then
                        Dim ss As String() = s.Split("|")
                        Dim cnum As String = ss(0).Trim
                        Dim mnum As String = ss(3).Trim
                        Dim newline As String = cnum & "," & mnum
                        If Not existingmatters.Contains(newline) Then
                            If cnum.Length > 4 Then
                                Dim MatterType As String = cnum.Substring(0, 2)
                                Select Case MatterType
                                    Case "RN", "LV", "NV"
                                        MatterType = "NV"
                                    Case "BH"
                                        MatterType = "BH"
                                    Case Else
                                        MatterType = "CA"
                                End Select
                                Dim aMatter As String = s.Replace("""", "") & MatterType
                                If Not newmatters.Contains(aMatter) Then
                                    newmatters.Add(aMatter)
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            Loop
        End Using
        Environment.ExitCode = newmatters.Count
        Dim emailbody As String = "Filtered file: " & downloadfile & vbCrLf & vbCrLf & "New Matters: " & vbCrLf & vbCrLf
        Using sw As New StreamWriter(filteredfile, False)
            sw.AutoFlush = True
            For Each s As String In newmatters
                emailbody &= s & vbCrLf
                sw.WriteLine(s)
            Next
        End Using
        File.Copy(downloadfile, downloadfile & ".original", True)
        File.Copy(filteredfile, downloadfile, True)
        Console.WriteLine(emailbody)
        EmailReport(emailbody, newmatters.Count)
    End Sub
    Sub EmailReport(rBody As String, mf As Integer)
        Dim msg As New MailMessage
        With msg
            .From = New MailAddress("DMSImport@lawfirm.com")
            .To.Add("fcanton@lawfirm.com")
            .Body = rBody
            If mf > 0 Then
                .Subject = "DMS Import file filtered"
            Else
                .Subject = "DMS Import file filtered - No new matters found"
            End If
        End With
        Dim msgcl As New SmtpClient("smtp1")
        With msgcl
            Try
                .Send(msg)
            Catch ex As Exception
                .Host = "smtp1"
                .Send(msg)
            End Try
        End With
    End Sub
End Module
