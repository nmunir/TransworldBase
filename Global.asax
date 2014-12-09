<%@ Application Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script RunAt="server">

    Sub Application_Start(ByVal Sender As Object, ByVal E As EventArgs)
        ' Code that runs on application startup
    End Sub

    Sub Application_End(ByVal Sender As Object, ByVal E As EventArgs)
        ' Code that runs on application shutdown
    End Sub

    Sub Application_Error(ByVal Sender As Object, ByVal E As EventArgs)
        ' RecordError()
    End Sub

    Sub Session_Start(ByVal Sender As Object, ByVal E As EventArgs)
        ' Code that runs when a new session is started
    End Sub

    Sub Session_OnEnd(ByVal Sender As Object, ByVal E As EventArgs)
        Logoff()
        Session.RemoveAll()
    End Sub

    Sub Logoff()
        Dim sConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
        Dim oConn As New SqlConnection(sConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Logoff", oConn)

        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try
    End Sub

    '    Sub RecordError()
    '        Dim sConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")
    '        Dim oConn As New SqlConnection(sConn)
    '        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Logoff", oConn)'''''''
    '
    '        oCmd.CommandType = CommandType.StoredProcedure
    '        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
    '        paramCustomerKey.Value = Session("CustomerKey")
    '        oCmd.Parameters.Add(paramCustomerKey)
    '        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
    '        paramUserKey.Value = Session("UserKey")
    '        oCmd.Parameters.Add(paramUserKey)
    '        oConn.Open()
    '        oCmd.ExecuteNonQuery()
    '    End Sub
</script>