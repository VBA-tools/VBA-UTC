Attribute VB_Name = "UtcConverter"
''
' VBA-UTC v0.0.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - ISO 8601 parse error
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private utc_Loaded As Boolean

Public utc_UtcOffsetMinutes As Long
Public utc_Dst As Boolean
Public utc_Override As Boolean

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @param {Date} utc_UtcDate
' @return {Date} Local date
' -------------------------------------- '
Public Function ParseUtc(utc_UtcDate As Date) As Date
    utc_LoadUtcOffsetAndDst
    ParseUtc = utc_UtcDate + VBA.TimeSerial(0, utc_UtcOffsetMinutes, 0) + VBA.IIf(utc_Dst, VBA.TimeSerial(1, 0, 0), 0)
End Function

''
' Convert local date to UTC date
'
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' -------------------------------------- '
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
     utc_LoadUtcOffsetAndDst
     ConvertToUtc = utc_LocalDate - VBA.TimeSerial(0, utc_UtcOffsetMinutes, 0) - VBA.IIf(utc_Dst, VBA.TimeSerial(1, 0, 0), 0)
End Function

''
' Parse ISO 8601 date string to local date
'
' @param {Date} utc_IsoString
' @return {Date} Local date
' -------------------------------------- '
Public Function ParseIso(utc_IsoString As String) As Date
    utc_LoadUtcOffsetAndDst
    
    On Error GoTo ErrorHandling
    
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date
    
    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
    
    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
            
            If utc_OffsetMinutes < 0 Then
                utc_Offset = -TimeSerial(0, utc_UtcOffsetMinutes, 0) + VBA.IIf(utc_Dst, TimeSerial(1, 0, 0), 0)
            Else
                utc_Offset = TimeSerial(0, utc_UtcOffsetMinutes, 0) + VBA.IIf(utc_Dst, TimeSerial(1, 0, 0), 0)
            End If
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If
            
            If utc_OffsetIndex > 0 Then
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")
                
                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), VBA.CInt(utc_OffsetParts(2)))
                End Select
                
                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If
        
        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), VBA.CInt(utc_TimeParts(2)))
        End Select
        
        ParseIso = ParseIso + utc_Offset
    End If
    
    Exit Function
    
ErrorHandling:
    
    Err.Raise 10011, "UtcConverter.ParseIso", "ISO 8601 parse error for " & utc_IsoString
End Function

''
' Convert local date to ISO 8601 string
'
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' -------------------------------------- '
Public Function ConvertToIso(utc_LocalDate As Date) As String
    utc_LoadUtcOffsetAndDst
    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Sub utc_LoadUtcOffsetAndDst()
    If Not utc_Loaded And Not utc_Override Then
        ' TODO
        utc_Loaded = True
    End If
End Sub
