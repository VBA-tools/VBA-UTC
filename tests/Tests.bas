Attribute VB_Name = "Tests"
Private OffsetMinutes As Long

' May 6, 2004 7:08:09 PM
Private Const LocalDate As Date = 38113.7973263889
Private Const LocalIso As String = "2004-05-06T19:08:09.000Z"


Public Sub Run(Optional FilePath As String = "", Optional Offset As String = "-0400")
    Dim Suite As New TestSuite
    Suite.Description = "vba-utc"

    Dim Reporter As New FileReporter
    Reporter.WriteTo FilePath
    Reporter.ListenTo Suite

    Dim Immediate As New ImmediateReporter
    Immediate.ListenTo Suite

    OffsetMinutes = CLng(Offset) / 100 * 60

    ShouldParseUTC Suite.Test("should parse UTC")
    ShouldConvertToUTC Suite.Test("should convert to UTC")
    ParseISO8601 Suite.Group("ParseIso")
    ShouldConvertToISO8601 Suite.Test("should convert to ISO 8601")
End Sub

Sub ShouldParseUTC(Test As TestCase)
    ' May 6, 2004 11:08:09 PM
    Dim UtcDate As Date
    UtcDate = LocalDate - OffsetMinutes / 60 / 24

    Test.IsEqual DateToString(UtcConverter.ParseUtc(UtcDate)), DateToString(LocalDate)
End Sub

Sub ShouldConvertToUTC(Test As TestCase)
    Dim UtcDate As Date
    UtcDate = LocalDate - OffsetMinutes / 60 / 24

    Test.IsEqual DateToString(UtcConverter.ConvertToUtc(LocalDate)), DateToString(UtcDate)
End Sub

Sub ParseISO8601(Suite As TestSuite)
    Dim UtcDate As Date
    Dim UtcIso As String
    Dim TZOffsetHours As Integer
    Dim TZOffsetMinutes As Integer
    Dim Offset As String

    UtcDate = LocalDate - OffsetMinutes / 60 / 24
    UtcIso = VBA.Format$(UtcDate, "yyyy-mm-ddTHH:mm:ss.000Z")

    TZOffsetHours = Int(-OffsetMinutes / 60)
    TZOffsetMinutes = -(OffsetMinutes + (TZOffsetHours * 60))

    With Suite.Test("should parse ISO 8601")
        .IsEqual DateToString(UtcConverter.ParseIso(UtcIso)), "2004-05-06T19:08:09"
    End With

    With Suite.Test("should parse ISO 8601 with offset")
        Offset = VBA.Right$("0" & TZOffsetHours, 2) & ":" & VBA.Right$("0" & (TZOffsetMinutes + 1), 2) & ":02"
        .IsEqual DateToString(UtcConverter.ParseIso("2004-05-06T19:07:07-" & Offset)), "2004-05-06T19:08:09"
    End With

    With Suite.Test("should parse ISO 8601 with varying time format")
        Offset = VBA.Right$("0" & TZOffsetHours, 2)
        .IsEqual DateToString(UtcConverter.ParseIso("2004-05-06T19-" & Offset)), "2004-05-06T19:00:00"

        Offset = VBA.Right$("0" & TZOffsetHours, 2) & ":" & VBA.Right$("0" & (TZOffsetMinutes + 1), 2)
        .IsEqual DateToString(UtcConverter.ParseIso("2004-05-06T19:07-" & Offset)), "2004-05-06T19:08:00"
        .IsEqual DateToString(UtcConverter.ParseIso("2004-05-06T12Z")), _
            DateToString(DateSerial(2004, 5, 6) + TimeSerial(12, 0, 0) + OffsetMinutes / 60 / 24)
        .IsEqual DateToString(UtcConverter.ParseIso("2004-05-06T12:08Z")), _
            DateToString(DateSerial(2004, 5, 6) + TimeSerial(12, 8, 0) + OffsetMinutes / 60 / 24)
    End With
End Sub

Sub ShouldConvertToISO8601(Test As TestCase)
    Dim UtcDate As Date
    Dim UtcIso As String

    UtcDate = LocalDate - OffsetMinutes / 60 / 24
    UtcIso = VBA.Format$(UtcDate, "yyyy-mm-ddTHH:mm:ss.000Z")

    Test.IsEqual UtcConverter.ConvertToIso(LocalDate), UtcIso
End Sub

Private Function DateToString(Value As Date) As String
    DateToString = VBA.Format$(Value, "yyyy-mm-ddTHH:mm:ss")
End Function
