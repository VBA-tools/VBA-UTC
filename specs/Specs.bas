Attribute VB_Name = "Specs"
Private pOffsetMinutes As Long
Private pOffsetLoaded As Boolean
Public Property Get OffSetMinutes() As Long
    If Not pOffsetLoaded Then
        Dim InputValue As String
        InputValue = VBA.InputBox("Enter UTC Offset (in minutes)" & vbNewLine & vbNewLine & _
            "Example:" & vbNewLine & _
            "EST (UTC-5:00) and DST (+1:00)" & vbNewLine & _
            "= UTC-4:00" & vbNewLine & _
            "= -240", "Enter UTC Offset", 0)
        
        If InputValue <> "" Then: pOffsetMinutes = CLng(InputValue)
        
        pOffsetLoaded = True
    End If
    
    OffSetMinutes = pOffsetMinutes
End Property
Public Property Let OffSetMinutes(Value As Long)
    pOffsetMinutes = Value
    pOffsetLoaded = True
End Property

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-UTC"
    
    Dim LocalDate As Date
    Dim LocalIso As String
    Dim UtcDate As Date
    Dim UtcIso As String
    
    ' May 6, 2004 7:08:09 PM
    LocalDate = 38113.7973263889
    LocalIso = "2004-05-06T19:08:09.000Z"
    
    ' May 6, 2004 11:08:09 PM
    UtcDate = LocalDate - OffSetMinutes / 60 / 24
    UtcIso = VBA.Format$(UtcDate, "yyyy-mm-ddTHH:mm:ss.000Z")
    
    ' ============================================= '
    ' ParseUTC
    ' ============================================= '
    With Specs.It("should parse UTC")
        .Expect(DateToString(UtcConverter.ParseUtc(UtcDate))).ToEqual DateToString(LocalDate)
    End With
    
    ' ============================================= '
    ' ConvertToUTC
    ' ============================================= '
    With Specs.It("should convert to UTC")
        .Expect(DateToString(UtcConverter.ConvertToUtc(LocalDate))).ToEqual DateToString(UtcDate)
    End With
    
    ' ============================================= '
    ' ParseISO
    ' ============================================= '
    With Specs.It("should parse ISO 8601")
        .Expect(DateToString(UtcConverter.ParseIso(UtcIso))).ToEqual "2004-05-06T19:08:09"
    End With
    
    With Specs.It("should parse ISO 8601 with offset")
        .Expect(DateToString(UtcConverter.ParseIso("2004-05-06T12:08:09+04:05:06"))).ToEqual "2004-05-06T16:13:15"
        .Expect(DateToString(UtcConverter.ParseIso("2004-05-06T12:08:09-04:05:06"))).ToEqual "2004-05-06T08:03:03"
    End With
    
    With Specs.It("should parse ISO 8601 with varying time format")
        .Expect(DateToString(UtcConverter.ParseIso("2004-05-06T12+04"))).ToEqual "2004-05-06T16:00:00"
        .Expect(DateToString(UtcConverter.ParseIso("2004-05-06T12:08+04:05"))).ToEqual "2004-05-06T16:13:00"
        .Expect(DateToString(UtcConverter.ParseIso("2004-05-06T12Z"))).ToEqual _
            DateToString(DateSerial(2004, 5, 6) + TimeSerial(12, 0, 0) + OffSetMinutes / 60 / 24)
        .Expect(DateToString(UtcConverter.ParseIso("2004-05-06T12:08Z"))).ToEqual _
            DateToString(DateSerial(2004, 5, 6) + TimeSerial(12, 8, 0) + OffSetMinutes / 60 / 24)
    End With
    
    ' ============================================= '
    ' ConvertToISO
    ' ============================================= '
    With Specs.It("should convert to ISO 8601")
        .Expect(UtcConverter.ConvertToIso(LocalDate)).ToEqual UtcIso
    End With
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    On Error Resume Next
    
    
    InlineRunner.RunSuite Specs
End Function

Public Sub RunSpecs()
    DisplayRunner.IdCol = 1
    DisplayRunner.DescCol = 1
    DisplayRunner.ResultCol = 2
    DisplayRunner.OutputStartRow = 4
    
    DisplayRunner.RunSuite Specs
End Sub

Private Function DateToString(Value As Date) As String
    DateToString = VBA.Format$(Value, "yyyy-mm-ddTHH:mm:ss")
End Function
