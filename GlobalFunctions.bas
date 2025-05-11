Attribute VB_Name = "GlobalFunctions"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
    
Public Function ArrayIsInitialized(arr) As Boolean
    Dim memVal As Long
    
    CopyMemory memVal, ByVal VarPtr(arr) + 8, ByVal 4 'get pointer to array
    CopyMemory memVal, ByVal memVal, ByVal 4 'see if it points to an address...
    ArrayIsInitialized = (memVal <> 0) '...if it does, array is initialized
End Function

