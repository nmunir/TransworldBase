Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography

Namespace SprintInternational
    Public Class Logon
        Public Function GetUserInfo(ByVal sUserId As String) As UserInfo
            Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString()
            Dim oConn As New SqlConnection(sConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_ValidateId5", oConn)
            oCmd.CommandType = CommandType.StoredProcedure

            Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.NVarChar, 100)
            paramUserId.Value = CStr(sUserId)
            oCmd.Parameters.Add(paramUserId)

            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramUserKey.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramUserKey)

            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            paramCustomerKey.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramCustomerKey)

            Dim paramCustomerName As SqlParameter = New SqlParameter("@CustomerName", SqlDbType.NVarChar, 50)
            paramCustomerName.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramCustomerName)

            Dim paramUserName As SqlParameter = New SqlParameter("@UserName", SqlDbType.NVarChar, 100)
            paramUserName.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramUserName)

            Dim paramPassword As SqlParameter = New SqlParameter("@Password", SqlDbType.NVarChar, 24)
            paramPassword.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramPassword)

            Dim paramUserType As SqlParameter = New SqlParameter("@Type", SqlDbType.NVarChar, 20)
            paramUserType.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramUserType)

            Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.NVarChar, 20)
            paramStatus.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramStatus)

            Dim paramViewGAB As SqlParameter = New SqlParameter("@AbleToViewGlobalAddressBook", SqlDbType.Bit, 1)
            paramViewGAB.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramViewGAB)

            Dim paramEditGAB As SqlParameter = New SqlParameter("@AbleToEditGlobalAddressBook", SqlDbType.Bit, 1)
            paramEditGAB.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramEditGAB)

            Dim paramCreateStockBooking As SqlParameter = New SqlParameter("@AbleToCreateStockBooking", SqlDbType.Bit, 1)
            paramCreateStockBooking.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramCreateStockBooking)

            Dim paramViewStock As SqlParameter = New SqlParameter("@AbleToViewStock", SqlDbType.Bit, 1)
            paramViewStock.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramViewStock)

            Dim paramCreateCollectionRequest As SqlParameter = New SqlParameter("@AbleToCreateCollectionRequest", SqlDbType.Bit, 1)
            paramCreateCollectionRequest.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramCreateCollectionRequest)

            Dim paramApplyMaxGrabRule As SqlParameter = New SqlParameter("@ApplyStockMaxGrabRule", SqlDbType.Bit, 1)
            paramApplyMaxGrabRule.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramApplyMaxGrabRule)

            Dim paramRunningHeader As SqlParameter = New SqlParameter("@RunningHeaderImage", SqlDbType.NVarChar, 100)
            paramRunningHeader.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramRunningHeader)

            Dim paramDefaultWebsite As SqlParameter = New SqlParameter("@DefaultWebsite", SqlDbType.NVarChar, 60)
            paramDefaultWebsite.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramDefaultWebsite)

            Dim paramMustChangePassword As SqlParameter = New SqlParameter("@MustChangePassword", SqlDbType.Bit, 1)
            paramMustChangePassword.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMustChangePassword)

            Dim paramLastPasswordChange As SqlParameter = New SqlParameter("@LastPasswordChange", SqlDbType.SmallDateTime)
            paramLastPasswordChange.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramLastPasswordChange)

            Dim paramMaxPasswordRetries As SqlParameter = New SqlParameter("@MaxPasswordRetries", SqlDbType.Int)
            paramMaxPasswordRetries.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMaxPasswordRetries)

            Dim paramMinPasswordLength As SqlParameter = New SqlParameter("@MinPasswordLength", SqlDbType.Int)
            paramMinPasswordLength.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMinPasswordLength)

            Dim paramMinPasswordLowerCaseChars As SqlParameter = New SqlParameter("@MinPasswordLowerCaseChars", SqlDbType.Int)
            paramMinPasswordLowerCaseChars.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMinPasswordLowerCaseChars)

            Dim paramMinPasswordDigits As SqlParameter = New SqlParameter("@MinPasswordDigits", SqlDbType.Int)
            paramMinPasswordDigits.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMinPasswordDigits)

            Dim paramPasswordExpiryDays As SqlParameter = New SqlParameter("@PasswordExpiryDays", SqlDbType.Int)
            paramPasswordExpiryDays.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramPasswordExpiryDays)

            Dim paramUserPermissions As SqlParameter = New SqlParameter("@UserPermissions", SqlDbType.Int)
            paramUserPermissions.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramUserPermissions)

            Dim paramLastLogon As SqlParameter = New SqlParameter("@LastLogon", SqlDbType.SmallDateTime)
            paramLastLogon.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramLastLogon)

            Dim paramAccountDisabledDueToInactivity As SqlParameter = New SqlParameter("@AccountDisabledDueToInactivity", SqlDbType.Bit, 1)
            paramAccountDisabledDueToInactivity.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramAccountDisabledDueToInactivity)

            Dim paramMemorableAnswer1 As SqlParameter = New SqlParameter("@MemorableAnswer1", SqlDbType.NVarChar, 50)
            paramMemorableAnswer1.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMemorableAnswer1)

            Dim paramMemorableAnswer2 As SqlParameter = New SqlParameter("@MemorableAnswer2", SqlDbType.NVarChar, 50)
            paramMemorableAnswer2.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMemorableAnswer2)

            Dim paramMemorableAnswer3 As SqlParameter = New SqlParameter("@MemorableAnswer3", SqlDbType.NVarChar, 50)
            paramMemorableAnswer3.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramMemorableAnswer3)

            oConn.Open()
            oCmd.ExecuteNonQuery()
            oConn.Close()

            Dim oUserInfo As UserInfo = New UserInfo()

            If Not IsDBNull(paramUserKey.Value) Then
                oUserInfo.UserKey = CInt(paramUserKey.Value)
            Else
                oUserInfo.UserKey = -1
            End If

            If Not IsDBNull(paramCustomerKey.Value) Then
                oUserInfo.CustomerKey = CInt(paramCustomerKey.Value)
            Else
                oUserInfo.CustomerKey = 0
            End If

            If Not IsDBNull(paramCustomerName.Value) Then
                oUserInfo.CustomerName = CStr(paramCustomerName.Value)
            Else
                oUserInfo.CustomerName = ""
            End If

            If Not IsDBNull(paramUserName.Value) Then
                oUserInfo.UserName = CStr(paramUserName.Value)
            Else
                oUserInfo.UserName = ""
            End If

            If Not IsDBNull(paramPassword.Value) Then
                oUserInfo.Password = CStr(paramPassword.Value)
            Else
                oUserInfo.Password = ""
            End If

            If Not IsDBNull(paramUserType.Value) Then
                oUserInfo.UserType = CStr(paramUserType.Value)
            Else
                oUserInfo.UserType = ""
            End If

            If Not IsDBNull(paramStatus.Value) Then
                oUserInfo.Status = CStr(paramStatus.Value)
            Else
                oUserInfo.Status = ""
            End If

            If Not IsDBNull(paramViewGAB.Value) Then
                oUserInfo.AbleToViewGlobalAddressBook = CBool(paramViewGAB.Value)
            Else
                oUserInfo.AbleToViewGlobalAddressBook = False
            End If

            If Not IsDBNull(paramEditGAB.Value) Then
                oUserInfo.AbleToEditGlobalAddressBook = CBool(paramEditGAB.Value)
            Else
                oUserInfo.AbleToEditGlobalAddressBook = False
            End If

            If Not IsDBNull(paramViewStock.Value) Then
                oUserInfo.AbleToViewStock = CBool(paramViewStock.Value)
            Else
                oUserInfo.AbleToViewStock = False
            End If

            If Not IsDBNull(paramCreateStockBooking.Value) Then
                oUserInfo.AbleToCreateStockBooking = CBool(paramCreateStockBooking.Value)
            Else
                oUserInfo.AbleToCreateStockBooking = False
            End If

            If Not IsDBNull(paramCreateCollectionRequest.Value) Then
                oUserInfo.AbleToCreateCollectionRequest = CBool(paramCreateCollectionRequest.Value)
            Else
                oUserInfo.AbleToCreateCollectionRequest = False
            End If

            If Not IsDBNull(paramApplyMaxGrabRule.Value) Then
                oUserInfo.ApplyStockMaxGrabRule = CBool(paramApplyMaxGrabRule.Value)
            Else
                oUserInfo.ApplyStockMaxGrabRule = False
            End If

            If Not IsDBNull(paramRunningHeader.Value) Then
                oUserInfo.RunningHeaderImage = CStr(paramRunningHeader.Value)
            Else
                oUserInfo.RunningHeaderImage = "default"
            End If

            If Not IsDBNull(paramDefaultWebsite.Value) Then
                oUserInfo.DefaultWebsite = CStr(paramDefaultWebsite.Value)
            Else
                oUserInfo.DefaultWebsite = ""
            End If

            If Not IsDBNull(paramMustChangePassword.Value) Then
                oUserInfo.MustChangePassword = CBool(paramMustChangePassword.Value)
            Else
                oUserInfo.MustChangePassword = False
            End If

            If Not IsDBNull(paramLastPasswordChange.Value) Then
                oUserInfo.LastPasswordChange = CDate(paramLastPasswordChange.Value)
            Else
                oUserInfo.LastPasswordChange = Now
            End If

            If Not IsDBNull(paramMaxPasswordRetries.Value) Then
                oUserInfo.MaxPasswordRetries = CInt(paramMaxPasswordRetries.Value)
            Else
                oUserInfo.MaxPasswordRetries = 9999
            End If

            If Not IsDBNull(paramMinPasswordLength.Value) Then
                oUserInfo.MinPasswordLength = CInt(paramMinPasswordLength.Value)
            Else
                oUserInfo.MinPasswordLength = 0
            End If

            If Not IsDBNull(paramMinPasswordLowerCaseChars.Value) Then
                oUserInfo.MinPasswordLowerCaseChars = CInt(paramMinPasswordLowerCaseChars.Value)
            Else
                oUserInfo.MinPasswordLowerCaseChars = 0
            End If

            If Not IsDBNull(paramMinPasswordDigits.Value) Then
                oUserInfo.MinPasswordDigits = CInt(paramMinPasswordDigits.Value)
            Else
                oUserInfo.MinPasswordDigits = 0
            End If

            If Not IsDBNull(paramPasswordExpiryDays.Value) Then
                oUserInfo.PasswordExpiryDays = CInt(paramPasswordExpiryDays.Value)
            Else
                oUserInfo.PasswordExpiryDays = 10000
            End If

            If Not IsDBNull(paramUserPermissions.Value) Then
                oUserInfo.UserPermissions = CInt(paramUserPermissions.Value)
            Else
                oUserInfo.UserPermissions = 0
            End If

            If Not IsDBNull(paramLastLogon.Value) Then
                oUserInfo.LastLogon = CDate(paramLastLogon.Value)
            Else
                oUserInfo.LastLogon = DateTime.MinValue
            End If

            If Not IsDBNull(paramAccountDisabledDueToInactivity.Value) Then
                oUserInfo.AccountDisabledDueToInactivity = CBool(paramAccountDisabledDueToInactivity.Value)
            Else
                oUserInfo.AccountDisabledDueToInactivity = False
            End If

            If Not IsDBNull(paramMemorableAnswer1.Value) Then
                oUserInfo.MemorableAnswer1 = CStr(paramMemorableAnswer1.Value)
            Else
                oUserInfo.MemorableAnswer1 = String.Empty
            End If

            If Not IsDBNull(paramMemorableAnswer2.Value) Then
                oUserInfo.MemorableAnswer2 = CStr(paramMemorableAnswer2.Value)
            Else
                oUserInfo.MemorableAnswer2 = String.Empty
            End If

            If Not IsDBNull(paramMemorableAnswer3.Value) Then
                oUserInfo.MemorableAnswer3 = CStr(paramMemorableAnswer3.Value)
            Else
                oUserInfo.MemorableAnswer3 = String.Empty
            End If

            Return oUserInfo
        End Function
    End Class

    ' UserPermissions values:
    ' 1 = Account Handler
    ' 2, 4 = Site & Deputy Site Administrator
    ' 8, 16 = Notice Board & Deputy Notice Board Editor

    Public Class UserInfo
        Public UserKey As Int32
        Public CustomerKey As Int32
        Public CustomerName As String
        Public UserName As String
        Public Password As String
        Public UserType As String
        Public Status As String
        Public AbleToViewGlobalAddressBook As Boolean
        Public AbleToEditGlobalAddressBook As Boolean
        Public AbleToViewStock As Boolean
        Public AbleToCreateStockBooking As Boolean
        Public AbleToCreateCollectionRequest As Boolean
        Public ApplyStockMaxGrabRule As Boolean
        Public RunningHeaderImage As String
        Public DefaultWebsite As String
        Public MustChangePassword As Boolean
        Public LastPasswordChange As DateTime
        Public MaxPasswordRetries As Int32
        Public MinPasswordLength As Int32
        Public MinPasswordLowerCaseChars As Int32
        Public MinPasswordDigits As Int32
        Public PasswordExpiryDays As Int32
        Public UserPermissions As Int32
        Public LastLogon As DateTime
        Public AccountDisabledDueToInactivity As Boolean
        Public MemorableAnswer1 As String
        Public MemorableAnswer2 As String
        Public MemorableAnswer3 As String
    End Class

    Public Class Password
        Const DESKey As String = "!#$a54?3"
        Const DESIV As String = "|~#:+=*&"
        Private key() As Byte
        Private IV() As Byte

        Public Function Decrypt(ByVal stringToDecrypt As String) As String
            Dim inputByteArray(stringToDecrypt.Length) As Byte
            Try
                key = System.Text.Encoding.UTF8.GetBytes(Left(DESKey, 8))
                IV = System.Text.Encoding.UTF8.GetBytes(Left(DESIV, 8))
                Dim des As New DESCryptoServiceProvider()
                inputByteArray = Convert.FromBase64String(stringToDecrypt)
                Dim ms As New MemoryStream()
                Dim cs As New CryptoStream(ms, des.CreateDecryptor(key, IV), _
                    CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
                Return encoding.GetString(ms.ToArray())
            Catch e As Exception
                Return e.Message
            End Try
        End Function

        Public Function Encrypt(ByVal stringToEncrypt As String) As String
            Try
                key = System.Text.Encoding.UTF8.GetBytes(Left(DESKey, 8))
                IV = System.Text.Encoding.UTF8.GetBytes(Left(DESIV, 8))
                Dim des As New DESCryptoServiceProvider()
                Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(stringToEncrypt)
                Dim ms As New MemoryStream()
                Dim cs As New CryptoStream(ms, des.CreateEncryptor(key, IV), _
                    CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                Return Convert.ToBase64String(ms.ToArray())
            Catch e As Exception
                Return e.Message
            End Try
        End Function
    End Class
End Namespace
