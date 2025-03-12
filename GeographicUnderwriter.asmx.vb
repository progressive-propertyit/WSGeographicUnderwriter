Imports System.Web.Services
Imports System.ComponentModel
Imports PropertyIT.Common.Data

<WebService(Namespace:="http://services.e-ins.net/WSGeographicUnderwriter/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class GeographicUnderwriter
    Inherits WebService

    Private Const NOTFOUNDZIP As String = "34769"
    Private Const MILE As Double = 0.016667
    Private Const MAXZIPCODEFACTOR As Double = 1.4
    Private Const NODISTANCESCORE As Integer = 12
    Private Const STATEDIVFACTOR As Integer = 2

    Private _DBase As SQLBase_DAL

    Private ErrorMessages As New List(Of String)

    Private ReadOnly Property DBase() As SQLBase_DAL
        Get
            If _DBase Is Nothing Then
                Dim HomeownersRO As String = ConfigurationManager.ConnectionStrings(NameOf(HomeownersRO)).ConnectionString
                _DBase = New SQLBase_DAL(HomeownersRO)
            End If

            Return _DBase
        End Get
    End Property

    Private Function GetZipCodeFactor(ZipCode As String) As Double
        Const SQL_DEMO_ZIPCODE As String = "SELECT HouseHolds FROM GeoDemographics WHERE ZipCode = @ZipCode"
        Const SQL_AVG_ZIPCODE As String = "SELECT Avg(HouseHolds) HouseHolds FROM GeoDemographics"

        Dim intHouseHolds As Integer,
            intAvgHouseHolds As Integer
        Dim zcf As Double
        Try
            intHouseHolds = 0
            intAvgHouseHolds = Convert.ToInt32(DBase.ExecScalar(SQL_AVG_ZIPCODE))

            If ZipCode = "" Then ZipCode = NOTFOUNDZIP

            Dim params As New List(Of SqlClient.SqlParameter)
            params.Add(New SqlClient.SqlParameter("@ZipCode", SqlDbType.VarChar, 5) With {.Value = ZipCode})
            Using dt As DataTable = DBase.OpenQuery(SQL_DEMO_ZIPCODE, params.ToArray)
                If dt.Rows.Count = 0 Then
                    intHouseHolds = 0
                Else
                    intHouseHolds = Integer.Parse(dt.Rows(0)("HouseHolds").ToString)
                End If
            End Using

            If intHouseHolds = 0 Then
                params(0).Value = NOTFOUNDZIP
                Using dt As DataTable = DBase.OpenQuery(SQL_DEMO_ZIPCODE, params.ToArray)
                    If dt.Rows.Count = 0 Then
                        intHouseHolds = 0
                    Else
                        intHouseHolds = Integer.Parse(dt.Rows(0)("HouseHolds").ToString)
                    End If
                End Using
            End If

            zcf = (1 + intHouseHolds / intAvgHouseHolds) / 2
            zcf = Math.Min(zcf, MAXZIPCODEFACTOR)

            Return Math.Round(zcf, 2, MidpointRounding.AwayFromZero)

        Catch ex As Exception
            ErrorMessages.Add(String.Format("Error caught in {0} - {1}", ex.Source, ex.Message))
            Return -1

        End Try

    End Function

    Private Function GetScore(theType As String, Latitude As Double, Longitude As Double) As Double
        Try
            Const SQL_DW As String = "SELECT Distance, Weight FROM GeoDWLookups WHERE TYPE = @Type"

            Dim Distance As Double,
                DWDistance As Double,
                DWWeight As Double,
                dblScore As Double = 0

            Dim params As New List(Of SqlClient.SqlParameter)
            params.Add(New SqlClient.SqlParameter("@Type", SqlDbType.VarChar, 25) With {.Value = theType})
            Using dt As DataTable = DBase.OpenQuery(SQL_DW, params.ToArray)
                DWDistance = Double.Parse(dt.Rows(0)("Distance").ToString)
                DWWeight = Double.Parse(dt.Rows(0)("Weight").ToString)
            End Using

            Using dt As DataTable = DBase.OpenQuery("Geo_SinkholeData", params.ToArray)
                For Each dr As DataRow In dt.Rows
                    Dim LatitudeDistance As Double = Double.Parse(dr("Longitude").ToString) - Longitude
                    Dim LongitudeDistance As Double = Double.Parse(dr("Latitude").ToString) - Latitude

                    Distance = Math.Sqrt(((Double.Parse(dr("Longitude").ToString) - Longitude) ^ 2) + ((Double.Parse(dr("Latitude").ToString) - Latitude) ^ 2)) / MILE

                    If Distance = 0 Then
                        dblScore += NODISTANCESCORE
                    ElseIf Distance < DWDistance Then
                        dblScore += (1 / Distance)
                    End If
                Next
            End Using

            Return (dblScore * DWWeight)

        Catch ex As Exception
            ErrorMessages.Add(String.Format("Error caught in {0} - {1}", ex.Source, ex.Message))
            Return -1

        End Try
    End Function

    Private Function GetWeightedScore(Latitude As Double, Longitude As Double) As Double
        Try
            Const SQL_TYPES As String = "SELECT distinct type, selectorder FROM GeoDWLookups ORDER BY SelectOrder"

            Dim Modifier As Integer
            Dim dblScore As Double

            Using dt As DataTable = DBase.OpenQuery(SQL_TYPES)
                For Each dr As DataRow In dt.Rows
                    Dim theType As String = dr("Type").ToString.ToUpper
                    If theType = "STATE" Then
                        If dblScore = 0 Then
                            Modifier = STATEDIVFACTOR
                        Else
                            Modifier = 1
                        End If
                    End If
                    dblScore += GetScore(theType, Latitude, Longitude)
                Next
            End Using

            Return (dblScore / Modifier)

        Catch ex As Exception
            ErrorMessages.Add(String.Format("Error caught in {0} - {1}", ex.Source, ex.Message))
            Return -1

        End Try
    End Function

    <WebMethod(Description:="Method to score Geo via ZipCode, Latitude and Longitude",
               EnableSession:=False)>
    Public Function ScoreAddressByLatLong(ZipCode As String, Longitude As Double, Latitude As Double, ByRef Errors As List(Of String)) As Double
        Try
            If ZipCode.Length <> 5 Then ZipCode = ZipCode.Substring(0, 5)

            Dim ZCFactor As Double = GetZipCodeFactor(ZipCode)
            Dim WeightedScore As Double = GetWeightedScore(Latitude, Longitude)

            If ErrorMessages.Count = 0 Then
                Return Math.Round(WeightedScore / ZCFactor, 2, MidpointRounding.AwayFromZero)
            Else
                Return -10
            End If

        Catch ex As Exception
            ErrorMessages.Add(String.Format("Error caught in {0} - {1}", ex.Source, ex.Message))
            Return -10

        Finally
            Errors = ErrorMessages
        End Try
    End Function

End Class