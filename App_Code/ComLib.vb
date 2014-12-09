Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class ComLib

    Private Shared gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private Shared goControl As Control = Nothing

    Public Shared Function ExecuteQueryToDataTable(ByVal sQuery As String, Optional ByVal oControl As Control = Nothing) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oConn.Open()
            If oControl IsNot Nothing Then
                goControl = oControl
                AddHandler oConn.InfoMessage, Function(sender, f) ExecuteQueryToDataTableAnonymousMethod(sender, f)
            End If
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
            Throw New Exception(ex.Message & vbCrLf & "Unexpected, catastrophic SQL error - processing terminated!")
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function

    Private Shared Function ExecuteQueryToDataTableAnonymousMethod(ByVal sender As Object, ByVal f As SqlInfoMessageEventArgs) As Boolean
        Dim tb As TextBox = Nothing
        Dim lbl As Label = Nothing
        Dim lb As ListBox = Nothing
        Dim ddl As DropDownList = Nothing
        If TypeOf (goControl) Is TextBox Then
            tb = goControl
            tb.Text += Constants.vbLf + f.Message
        ElseIf TypeOf goControl Is Label Then
            lbl = goControl
            lbl.Text += "<br />" + f.Message
        ElseIf TypeOf goControl Is ListBox Then
            lb = goControl
            lb.Items.Add(f.Message)
        ElseIf TypeOf goControl Is DropDownList Then
            ddl = goControl
            ddl.Items.Add(f.Message)
        End If
        Return True
    End Function

End Class
