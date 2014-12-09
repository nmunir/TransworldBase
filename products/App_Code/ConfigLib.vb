Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.Configuration.ConfigurationManager

Public Class ConfigLib
    Public Shared Function GetConfigItem_ConnectionString() As String
        GetConfigItem_ConnectionString = ""
        Try
            GetConfigItem_ConnectionString = ConnectionStrings("TempConnectionString").ConnectionString
        Catch e As System.NullReferenceException
            Try
                GetConfigItem_ConnectionString = ConnectionStrings("AIMSRootConnectionString").ConnectionString
            Catch e2 As System.NullReferenceException
                WebMsgBox.Show("No database connection string defined! (TempConnectionString / AIMSRootConnectionString)")
            End Try
        End Try
    End Function

    Public Shared Function GetConfigItem_AppTitle() As String
        Try
            GetConfigItem_AppTitle = AppSettings.Item("AppTitle")
        Catch e As System.NullReferenceException
            GetConfigItem_AppTitle = ""
        End Try
    End Function
End Class
