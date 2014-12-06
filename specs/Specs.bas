Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-UtcConverter"
    
    ' Override UTC Offset and DST for specs
    UtcConverter.utc_UtcOffsetMinutes = -5 * 60
    UtcConverter.utc_Dst = True
    UtcConverter.utc_Override = True
    
    Dim LocalDate As Date
    Dim LocalIso As String
    Dim UtcDate As Date
    Dim UtcIso As String
    
    ' May 6, 2004 7:08:09 PM
    LocalDate = 38113.7973263889
    LocalIso = "2004-05-06T19:08:09.000Z"
    
    ' May 6, 2004 11:08:09 PM
    UtcDate = 38113.9639930556
    UtcIso = "2004-05-06T23:08:09.000Z"
    
    ' ============================================= '
    ' ParseUTC
    ' ============================================= '
    With Specs.It("should parse UTC by UTCOffset and DST")
        .Expect(VBA.Format$(UtcConverter.ParseUtc(UtcDate), "yyyy-mm-ddTHH:mm:ss.000Z")).ToEqual LocalIso
    End With
    
    ' ============================================= '
    ' ConvertToUTC
    ' ============================================= '
    With Specs.It("should convert to UTC by UTCOffset and DST")
        .Expect(VBA.Format$(UtcConverter.ConvertToUtc(LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")).ToEqual UtcIso
    End With
    
    ' ============================================= '
    ' ParseISO
    ' ============================================= '
    With Specs.It("should parse ISO 8601 by UTCOffset and DST")
        .Expect(VBA.Format$(UtcConverter.ParseIso(UtcIso), "yyyy-mm-ddTHH:mm:ss.000Z")).ToEqual LocalIso
    End With
    
    With Specs.It("should parse ISO 8601 with offset")
        .Expect(VBA.Format$(UtcConverter.ParseIso("2004-05-06T12:08:09+04:05:06"), "yyyy-mm-ddTHH:mm:ss")).ToEqual "2004-05-06T16:13:15"
        .Expect(VBA.Format$(UtcConverter.ParseIso("2004-05-06T12:08:09-04:05:06"), "yyyy-mm-ddTHH:mm:ss")).ToEqual "2004-05-06T08:03:03"
    End With
    
    With Specs.It("should parse ISO 8601 with varying time format")
        .Expect(VBA.Format$(UtcConverter.ParseIso("2004-05-06T12+04"), "yyyy-mm-ddTHH:mm:ss")).ToEqual "2004-05-06T16:00:00"
        .Expect(VBA.Format$(UtcConverter.ParseIso("2004-05-06T12:08+04:05"), "yyyy-mm-ddTHH:mm:ss")).ToEqual "2004-05-06T16:13:00"
        .Expect(VBA.Format$(UtcConverter.ParseIso("2004-05-06T12Z"), "yyyy-mm-ddTHH:mm:ss")).ToEqual "2004-05-06T08:00:00"
        .Expect(VBA.Format$(UtcConverter.ParseIso("2004-05-06T12:08Z"), "yyyy-mm-ddTHH:mm:ss")).ToEqual "2004-05-06T08:08:00"
    End With
    
    ' ============================================= '
    ' ConvertToISO
    ' ============================================= '
    With Specs.It("should convert to ISO 8601 by UTCOffset and DST")
        .Expect(VBA.Format$(UtcConverter.ConvertToIso(LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")).ToEqual UtcIso
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
