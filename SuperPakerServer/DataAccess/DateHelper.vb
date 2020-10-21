Namespace DataAccess
    Public Class DateHelper
        Private Shared ReadOnly FIRST_GOOD_DATE As New DateTime(1900, 1, 1)
        Public Shared Function MapDateLessThan1900ToNull(ByVal inputDate As System.Nullable(Of DateTime)) As System.Nullable(Of DateTime)
            Dim returnDate As System.Nullable(Of DateTime) = Nothing
            If inputDate >= FIRST_GOOD_DATE Then
                returnDate = inputDate
            End If
            Return returnDate
        End Function
    End Class
    '<TestFixture()> _
    'Public Class As_A_DateHelper
    '    <SetUp()> _
    '    Public Sub I_want_to()
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_null_date_when_passed_a_null_nullable_date()
    '        Dim testDate As System.Nullable(Of DateTime) = Nothing
    '        Assert.AreEqual(Nothing, DateHelper.MapDateLessThan1900ToNull(testDate))
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_null_date_when_passed_null()
    '        Assert.AreEqual(Nothing, DateHelper.MapDateLessThan1900ToNull(Nothing))
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_null_date_when_passed_Dec_31_1899()
    '        Dim testDate As New DateTime(1899, 12, 31)
    '        Assert.AreEqual(Nothing, DateHelper.MapDateLessThan1900ToNull(testDate))
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_date_when_passed_Jan_01_1900()
    '        Dim testDate As New DateTime(1900, 1, 1)
    '        Assert.AreEqual(testDate, DateHelper.MapDateLessThan1900ToNull(testDate))
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_date_when_passed_normal_date()
    '        Dim testDate As New DateTime(2008, 8, 6)
    '        Assert.AreEqual(testDate, DateHelper.MapDateLessThan1900ToNull(testDate))
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_null_date_when_passed_nullable_Dec_31_1899()
    '        Dim testDate As System.Nullable(Of DateTime) = New DateTime(1899, 12, 31)
    '        Assert.AreEqual(Nothing, DateHelper.MapDateLessThan1900ToNull(testDate))
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_date_when_passed_nullable_Jan_01_1900()
    '        Dim testDate As System.Nullable(Of DateTime) = New DateTime(1900, 1, 1)
    '        Assert.AreEqual(testDate, DateHelper.MapDateLessThan1900ToNull(testDate))
    '    End Sub
    '    <Test()> _
    '    Public Sub Verify_MapDateLessThan1900ToNull_returns_date_when_passed_normal_nullable_date()
    '        Dim testDate As System.Nullable(Of DateTime) = New DateTime(2008, 8, 6)
    '        Assert.AreEqual(testDate, DateHelper.MapDateLessThan1900ToNull(testDate))
    '    End Sub
    'End Class
End Namespace
