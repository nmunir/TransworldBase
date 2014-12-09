Imports Microsoft.VisualBasic

Public Structure AuthorisationInfo
    Dim sID As String
    Dim bIsAuthorisable As Boolean
    Dim bAuthorisationRecordFound As Boolean
    Dim sAvailableAuthorisation As String
    Dim sPendingAuthorisation As String
    Dim dtAuthorisationExpiryDateTime As DateTime
    Dim bAuthorisationExpired As Boolean
    Dim nAuthoriser As Int32
End Structure

