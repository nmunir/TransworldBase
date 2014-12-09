
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Configuration
Imports Microsoft.VisualBasic
Imports System.Text


Namespace Tariff

    Public Class CostEstimate

        Public WeightCharge As Double
        Public EstimatedPackagingWeight As Double
        Public NonDoCSurCharge As Double
        Public DiscountRate As Double
        Public LocalTaxRate As Double

    End Class

    Public Class CostCalculator

        Public Function GetCostEstimate(ByVal lCustomerKey As Long, _
                                        ByVal lServiceLevelKey As Long, _
                                        ByVal sDocumentFlag As String, _
                                        ByVal sEstimatePackagingFlag As String, _
                                        ByVal lCountryKey As Long, _
                                        ByVal sTown As String, _
                                        ByVal sPostCode As String, _
                                        ByVal dblWeight As Double) As CostEstimate


            Dim dblWeightCharge As Double
            Dim dblMatrixBandFee As Double
            Dim bIsBaseRate As Boolean = True
            Dim dblBaseRate As Double
            Dim dblRemainder As Double
            Dim dblProductWeight As Double = dblWeight
            Dim dblPackagingWeight As Double = 0.00

            ' Create CustomerDetails Struct
            Dim oCostEstimate As CostEstimate = New CostEstimate()

            Dim dr as DataRow
            Dim sConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
            Dim oConn As New SqlConnection(sConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spStockMngr_Tariff_GetZoneMatrixFromAddress",oConn)
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            Try
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = lCustomerKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ServiceLevelKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@ServiceLevelKey").Value = lServiceLevelKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DocumentFlag", SqlDbType.NVarChar, 1))
                oAdapter.SelectCommand.Parameters("@DocumentFlag").Value = sDocumentFlag
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CountryKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CountryKey").Value = lCountryKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Town", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@Town").Value = sTown
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@PostalCode", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@PostalCode").Value = sPostCode

                oAdapter.Fill(oDataSet, "ZoneMatrix")

                If sEstimatePackagingFlag = "Y" Then
                    Do While dblProductWeight > 0
                        'Add 230 grams (for packaging) per 12.5 kilos of product
                        dblPackagingWeight = dblPackagingWeight + 0.23
                        dblProductWeight = dblProductWeight - 12.5
                        'Now add packaging weight onto the consignment Weight
                        dblWeight = dblWeight + dblPackagingWeight
                    Loop
                End IF


                For Each dr in oDataSet.Tables("ZoneMatrix").Rows
                    'First make a record of the Base Charge
                    If bIsBaseRate Then
                        dblBaseRate = ((dr("WeightTo") - dr("WeightFrom")) / dr("Units")) * dr("Fee")
                        bIsBaseRate = False
                    End If
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Iterating through the zone matrix ~~~~~~~~~~~~~~~~~~~~~~~~
                    'If parcel weight is heavier than this row's from and to delimiters then work out how much
                    'this band charges out at, add it to the running total [dblWeightCharge] and go to next row.
                    If dblWeight >= dr("WeightTo") Then
                        'Normally the weight charge is calculated like your tax and you work out what each
                        'portion of the weight is charged at. This model is broken when the Tariff is a flat rate
                        'tariff. Flat Rate Tariffs just multiply the unit charge by the parcel weight divided by
                        'the units
                        If dr("FlatRate") = False Then  ' not a flat rate
                            dblMatrixBandFee = ((dr("WeightTo") - dr("WeightFrom")) / dr("Units")) * dr("Fee")
                            dblWeightCharge = dblWeightCharge + dblMatrixBandFee
                        Else                       ' this is now (possible already was) a flat rate charge
                            'This matrix row is marked as Flat Rate which means we now apply the this rows's unit
                            'charge to all units. Before disregarding everything we must also see if the 'Hold Base'
                            'flag is set. If it is we must still charge the first rows unit charge and only apply
                            'flat rate for units thereafter.
                            If dr("HoldBase") = True Then
                                dblWeightCharge = ((dblWeight / dr("Units")) * dr("Fee")) + dblBaseRate
                            Else
                                dblWeightCharge = (dblWeight / dr("Units")) * dr("Fee")
                            End If
                        End If
                    '~~~~~~~~~~~~~~~~~~~~~~ Stop here: weight lies between this row's from and to ~~~~~~~~~~~~~~~~~~~
                    'Else if parcel weight lies between this row's from and to delimiters then this is the last
                    'row we need look at in this Zone Matrix. Calculate the weight charge and add it to running total.
                    ElseIf dblWeight >= dr("WeightFrom") And dblWeight < dr("WeightTo") Then
                        If dr("FlatRate") = False Then  ' not a flat rate
                            dblRemainder = (dblWeight - dr("WeightFrom")) / dr("Units")
                            Do While dblRemainder > 0
                                dblWeightCharge = dblWeightCharge + dr("Fee")
                                dblRemainder = dblRemainder - 1
                            Loop
                        Else
                            If dr("HoldBase") = True Then  'see above for explanation
                                dblWeightCharge = ((dblWeight / dr("Units")) * dr("Fee")) + dblBaseRate
                            Else
                                dblWeightCharge = (dblWeight / dr("Units")) * dr("Fee")
                            End If
                        End If
                    End If
                    'Following variables returned in each row of recordset - all the same
                    'When I find out how to return more than one recordset from a stored procedure
                    'and then iterate through them, I'll change this code.
                    oCostEstimate.NonDoCSurCharge = CDbl(dr("NonDocSurcharge"))
                    oCostEstimate.DiscountRate = CDbl(dr("DiscountRate"))
                    oCostEstimate.LocalTaxRate = CDbl(dr("LocalTaxRate"))
                Next

                oCostEstimate.WeightCharge = dblWeightCharge
                oCostEstimate.EstimatedPackagingWeight = dblPackagingWeight

            Catch ex As SqlException
            Finally
                oConn.Close()
            End Try

            Return oCostEstimate

        End Function

	End Class

End Namespace
