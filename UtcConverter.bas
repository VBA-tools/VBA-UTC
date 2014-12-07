Attribute VB_Name = "UtcConverter"
''
' VBA-UTC v0.5.0
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

#If Mac Then
#Else
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

Private Type utc_SYSTEMTIME
    utc_A_wYear As Integer
    utc_B_wMonth As Integer
    utc_C_wDayOfWeek As Integer
    utc_D_wDay As Integer
    utc_E_wHour As Integer
    utc_F_wMinute As Integer
    utc_G_wSecond As Integer
    utc_H_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_A_Bias As Long
    utc_B_StandardName(0 To 31) As Integer
    utc_C_StandardDate As utc_SYSTEMTIME
    utc_D_StandardBias As Long
    utc_E_DaylightName(0 To 31) As Integer
    utc_F_DaylightDate As utc_SYSTEMTIME
    utc_G_DaylightBias As Long
End Type
#End If

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
#If Mac Then
    ' TODO
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME
    
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, DateToSystemTime(utc_UtcDate), utc_LocalDate
    
    ParseUtc = SystemTimeToDate(utc_LocalDate)
#End If
End Function

''
' Convert local date to UTC date
'
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' -------------------------------------- '
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
#If Mac Then
    ' TODO
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME
    
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, DateToSystemTime(utc_LocalDate), utc_UtcDate
    
    ConvertToUtc = SystemTimeToDate(utc_UtcDate)
#End If
End Function

''
' Parse ISO 8601 date string to local date
'
' @param {Date} utc_IsoString
' @return {Date} Local date
' -------------------------------------- '
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo ErrorHandling
    
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date
    
    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
    
    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If
            
            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
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
        
        If utc_HasOffset Then
            ParseIso = ParseIso + utc_Offset
        Else
            ParseIso = ParseUtc(ParseIso)
        End If
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
    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then
#Else
Private Function DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    DateToSystemTime.utc_A_wYear = VBA.Year(utc_Value)
    DateToSystemTime.utc_B_wMonth = VBA.Month(utc_Value)
    DateToSystemTime.utc_D_wDay = VBA.Day(utc_Value)
    DateToSystemTime.utc_E_wHour = VBA.Hour(utc_Value)
    DateToSystemTime.utc_F_wMinute = VBA.Minute(utc_Value)
    DateToSystemTime.utc_G_wSecond = VBA.Second(utc_Value)
    DateToSystemTime.utc_H_wMilliseconds = 0
End Function

Private Function SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    SystemTimeToDate = DateSerial(utc_Value.utc_A_wYear, utc_Value.utc_B_wMonth, utc_Value.utc_D_wDay) + _
        TimeSerial(utc_Value.utc_E_wHour, utc_Value.utc_F_wMinute, utc_Value.utc_G_wSecond)
End Function
#End If
