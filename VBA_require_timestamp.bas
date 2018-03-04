Attribute VB_Name = "VBA_require_timestamp"
Option Explicit

'Returns current timestamp in "yyyy-MM-dd hh:mm:ss" format
Function timestamp(Optional strFormat As String = "yyyy-MM-dd hh:mm:ss")
    timestamp = Format(Now(), strFormat)
End Function
