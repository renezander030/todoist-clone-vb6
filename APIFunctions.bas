Attribute VB_Name = "APIFunctions"
DefLng A-Z

Const INTERNET_INVALID_PORT_NUMBER = 0
Const INTERNET_OPEN_TYPE_PRECONFIG = 0

' Number of the TCP/IP port on the server to connect to.
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_DEFAULT_GOPHER_PORT = 70
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443
Private Const INTERNET_DEFAULT_SOCKS_PORT = 1080

' Type of service to access.
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3

' Brings the data across the wire even if it locally cached.
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_SECURE = &H800000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_ASYNC = &H10000000
Private Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100

'internet error flags
Private Const INTERNET_ERROR_BASE = 12000
Private Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Private Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Private Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Private Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Private Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Private Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Private Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Private Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Private Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Private Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Private Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Private Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Private Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Private Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Private Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Private Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Private Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Private Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Private Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Private Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Private Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Private Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Private Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Private Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Private Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Private Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Private Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Private Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Private Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Private Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Private Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Private Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)
Private Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Private Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Private Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Private Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Private Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Private Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Private Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Private Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Private Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Private Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Private Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Private Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Private Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Private Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Private Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Private Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Private Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)
Private Const ERROR_INTERNET_FORTEZZA_LOGIN_NEEDED = (INTERNET_ERROR_BASE + 54)
Private Const ERROR_INTERNET_SEC_CERT_ERRORS = (INTERNET_ERROR_BASE + 55)
Private Const ERROR_INTERNET_SEC_CERT_NO_REV = (INTERNET_ERROR_BASE + 56)
Private Const ERROR_INTERNET_SEC_CERT_REV_FAILED = (INTERNET_ERROR_BASE + 57)

Private Const INTERNET_OPTION_SECURITY_FLAGS As Long = 31

Private Const SECURITY_FLAG_IGNORE_REVOCATION = &H80
Private Const SECURITY_FLAG_IGNORE_WRONG_USAGE = &H200
Private Const SECURITY_FLAG_IGNORE_CERT_CN_INVALID = &H1000       'INTERNET_FLAG_IGNORE_CERT_CN_INVALID
Private Const SECURITY_FLAG_IGNORE_CERT_DATE_INVALID = &H2000       'INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
Private Const SECURITY_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000       'INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS
Private Const SECURITY_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000&       'INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP
Private Const SECURITY_FLAG_IGNORE_WEAK_SIGNATURE = &H10000
Private Const SECURITY_SET_MASK = (SECURITY_FLAG_IGNORE_REVOCATION Or SECURITY_FLAG_IGNORE_UNKNOWN_CA Or SECURITY_FLAG_IGNORE_WRONG_USAGE Or SECURITY_FLAG_IGNORE_CERT_CN_INVALID Or SECURITY_FLAG_IGNORE_CERT_DATE_INVALID Or SECURITY_FLAG_IGNORE_WEAK_SIGNATURE)

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
        (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
        ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    
Private Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByRef sBuffer As Any, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Long

Private Declare Function HttpSendRequest Lib _
"wininet.dll" Alias "HttpSendRequestA" _
(ByVal hHttpRequest As Long, ByVal sHeaders _
As String, ByVal lHeadersLength As Long, _
ByVal sOptional As String, ByVal _
lOptionalLength As Long) As Integer

Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Boolean


Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
   (ByVal hInet As Long) As Integer

Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByRef lpMultiByteStr As Byte, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 = 65001

Dim Body As String

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Private Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000


Function UTF8ToString(ByRef arr() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim sResult As String
    
    On Error GoTo Fail
    
    nBytes = UBound(arr) - LBound(arr) + 1
    If nBytes <= 0 Then Exit Function
    
    ' Get the required size for the output buffer
    nChars = MultiByteToWideChar(CP_UTF8, 0, arr(LBound(arr)), nBytes, 0, 0)
    If nChars <= 0 Then Exit Function
    
    ' Allocate buffer and convert
    sResult = String$(nChars, vbNullChar)
    MultiByteToWideChar CP_UTF8, 0, arr(LBound(arr)), nBytes, StrPtr(sResult), nChars
    
    UTF8ToString = sResult
    Exit Function
    
Fail:
    UTF8ToString = ""
    ' Safe fallback
End Function

Public Function GetTasks() As Collection
  
   Dim var1
   Dim hOpen As Long, hConnection As Long, hSet As Long, hHeader As Integer
   Dim hURL As Long
   Dim buffer(0 To 4095) As Byte
   Dim lNumBytesToRead As Long
   Dim bytesRead As Long
   Dim chunk() As Byte
   Dim allBytes() As Byte
   Dim totalLen As Long
   Dim firstRead As Boolean
   Dim offset As Long
   Dim responseBytes() As Byte
   Dim decodedText As String
   Dim result As String
   
   firstRead = True
   totalLen = 0
   
   Dim CGI As String
   
   CGI = "/api/v1/tasks/filter?query=today%7Coverdue"
   
   Body = ""
   
   Dim sHeaders                                As String
   sHeaders = "Authorization: Bearer " & APITOKEN & vbNewLine _
    & "Host: todoist.com" & vbNewLine _
    & "User-Agent: PostmanRuntime/7.43.3" & vbNewLine _
    & "Accept: */*" & vbNewLine _
    & "Postman-Token: 12345" & vbNewLine _
    & "Accept-Encoding: br" & vbNewLine _
    & "Connection: keep-alive" & vbNewLine _
    & "Cookie: csrf=abcd123" & vbNewLine
    
    On Error Resume Next
    
    Dim newHost As String

   ' open internet connection
      hOpen = InternetOpen("PostmanRuntime/7.43.3", INTERNET_OPEN_TYPE_PRECONFIG, "", "", 0)
   'hOpen = InternetOpen("PostmanRuntime/7.43.3", INTERNET_OPEN_TYPE_PRECONFIG, "", "", INTERNET_FLAG_ASYNC)
   If hOpen <> 0 Then
      hConnection = InternetConnect(hOpen, PROXY, INTERNET_DEFAULT_HTTP_PORT, "", "", INTERNET_SERVICE_HTTP, 0, 0)
      If hConnection <> 0 Then
      
         hURL = HttpOpenRequest(hConnection, "GET", CGI, 0, 0, 0, INTERNET_FLAG_RELOAD, 0)
         
         'MsgBox hURL
         If hURL <> 0 Then

            If HttpSendRequest(hURL, sHeaders, Len(sHeaders), Body, Len(Body)) Then
               
               Dim totalData() As Byte
               ReDim totalData(0)
               
               Do
                If InternetReadFile(hURL, buffer(0), 4096, bytesRead) = 0 Then Exit Do
                If bytesRead = 0 Then Exit Do
                
                ReDim chunk(0 To bytesRead - 1)
                Dim i As Long
                For i = 0 To bytesRead - 1
                    chunk(i) = buffer(i)
                Next
                
                If firstRead Then
                    allBytes = chunk
                    firstRead = False
                Else
                    Dim newLen As Long
                    newLen = totalLen + bytesRead
                    ReDim Preserve allBytes(0 To newLen - 1)
                    For i = 0 To bytesRead - 1
                        allBytes(totalLen + i) = buffer(i)
                    Next
                End If
                
                totalLen = totalLen + bytesRead
                
               Loop
                
                Dim p As Object
                decodedText = UTF8ToString(allBytes)
                Set p = JsonParseObject(decodedText)
                
                Dim jCollection As Collection
                Set jCollection = p("results")
                
                Dim tCollection As New Collection
                Dim t As Task, td As TaskDue
                Dim ItemIndex As Integer
            
                            
                For ItemIndex = 0 To jCollection.Count
                    Set t = New Task
                    Set td = New TaskDue
                    t.id = JsonValue(p, "results/" & ItemIndex & "/id")
                    t.Content = JsonValue(p, "results/" & ItemIndex & "/content")
                    t.Description = JsonValue(p, "results/" & ItemIndex & "/description")
                    t.priority = JsonValue(p, "results/" & ItemIndex & "/priority")
                    td.dueDate = JsonValue(p, "results/" & ItemIndex & "/due/date")
                    Set t.due = td
                    tCollection.Add t
               Next

            Else
               MsgBox "Error: " & Err.LastDllError
               MsgBox "Check API Token and Proxy server"
               GoTo ErrHandler
            End If
         Else
            MsgBox "Error 505"
            MsgBox "Check API Token and Proxy server"
            GoTo ErrHandler
         End If
      Else
         MsgBox "Error 506"
         MsgBox "Check API Token and Proxy server"
         GoTo ErrHandler
      End If
   Else
      MsgBox "Error 507"
      MsgBox "Check API Token and Proxy server"
      GoTo ErrHandler
   End If
   
   InternetCloseHandle hURL
   InternetCloseHandle hConnection
   InternetCloseHandle hOpen
   
    Set GetTasks = tCollection
GoTo EndFunction

ErrHandler:
    Set GetTasks = Nothing
    MsgBox "Error: " & Err.Description, vbExclamation
EndFunction:
End Function

Public Function AddNewTask(TaskContent As String, taskDueDate As String, priority As Integer)
  
   Dim var1
   Dim hOpen As Long, hConnection As Long, hSet As Long, hHeader As Integer
   Dim hURL As Long
   Dim sBuffer As String
   Dim lNumBytesToRead As Long
   Dim lNumberOfBytesRead As Long
   Dim result As String
   Dim buffer(0 To 4095) As Byte
   Dim bytesRead As Long
   Dim chunk() As Byte
   Dim allBytes() As Byte
   Dim totalLen As Long
   Dim firstRead As Boolean
   Dim offset As Long
   Dim responseBytes() As Byte
   Dim decodedText As String
   
   Dim CGI As String
   
   CGI = "/api/v1/tasks"
   
   Body = "{ ""content"": """ & TaskContent & """, ""due_string"": """ & taskDueDate & """, ""priority"":  """ & priority & """ }"
   
   Dim sHeaders As String
   sHeaders = "Authorization: Bearer " & APITOKEN & vbNewLine _
    & "Host: todoist.com" & vbNewLine _
    & "content-type: application/json" & vbNewLine _
    & "User-Agent: PostmanRuntime/7.43.3" & vbNewLine _
    & "Accept: */*" & vbNewLine _
    & "Postman-Token: 12345" & vbNewLine _
    & "Accept-Encoding: br" & vbNewLine _
    & "Connection: keep-alive" & vbNewLine _
    & "Cookie: csrf=abcd123" & vbNewLine
    
    Dim newHost As String

   ' open internet connection
      hOpen = InternetOpen("PostmanRuntime/7.43.3", INTERNET_OPEN_TYPE_PRECONFIG, "", "", 0)
   'hOpen = InternetOpen("PostmanRuntime/7.43.3", INTERNET_OPEN_TYPE_PRECONFIG, "", "", INTERNET_FLAG_ASYNC)
   If hOpen <> 0 Then
      hConnection = InternetConnect(hOpen, PROXY, INTERNET_DEFAULT_HTTP_PORT, "", "", INTERNET_SERVICE_HTTP, 0, 0)
      If hConnection <> 0 Then
      
         hURL = HttpOpenRequest(hConnection, "POST", CGI, 0, 0, 0, INTERNET_FLAG_RELOAD, 0)
         
         'MsgBox hURL
         If hURL <> 0 Then

            If HttpSendRequest(hURL, sHeaders, Len(sHeaders), Body, Len(Body)) Then

               Dim totalData() As Byte
               ReDim totalData(0)
               
               Do
                If InternetReadFile(hURL, buffer(0), 4096, bytesRead) = 0 Then Exit Do
                If bytesRead = 0 Then Exit Do
                
                ReDim chunk(0 To bytesRead - 1)
                Dim i As Long
                For i = 0 To bytesRead - 1
                    chunk(i) = buffer(i)
                Next
                
                If firstRead Then
                    allBytes = chunk
                    firstRead = False
                Else
                    Dim newLen As Long
                    newLen = totalLen + bytesRead
                    ReDim Preserve allBytes(0 To newLen - 1)
                    For i = 0 To bytesRead - 1
                        allBytes(totalLen + i) = buffer(i)
                    Next
                End If
                
                totalLen = totalLen + bytesRead
                
               Loop
            Else
               Debug.Print (Err.LastDllError)
               'syncState.Caption = SYNC_CAPTION_PREFIX & Err.LastDllError
            End If
         Else
            Debug.Print ("505")
         End If
      Else
         Debug.Print ("506")
      End If
   Else
      Debug.Print ("507")
   End If
   
   InternetCloseHandle hURL
   InternetCloseHandle hConnection
   InternetCloseHandle hOpen

End Function

Public Function CloseTask(taskId As String)
  
   Dim var1
   Dim hOpen As Long, hConnection As Long, hSet As Long, hHeader As Integer
   Dim hURL As Long
   Dim sBuffer As String
   Dim lNumBytesToRead As Long
   Dim lNumberOfBytesRead As Long
   Dim result As String
   Dim buffer(0 To 4095) As Byte
   Dim bytesRead As Long
   Dim chunk() As Byte
   Dim allBytes() As Byte
   Dim totalLen As Long
   Dim firstRead As Boolean
   Dim offset As Long
   Dim responseBytes() As Byte
   Dim decodedText As String
   
   Dim CGI As String
   
   CGI = "/api/v1/tasks/" & taskId & "/close"
   
   Body = ""
   
   Dim sHeaders As String
   sHeaders = "Authorization: Bearer " & APITOKEN & vbNewLine _
    & "Host: todoist.com" & vbNewLine _
    & "content-type: application/json" & vbNewLine _
    & "User-Agent: PostmanRuntime/7.43.3" & vbNewLine _
    & "Accept: */*" & vbNewLine _
    & "Postman-Token: 12345" & vbNewLine _
    & "Accept-Encoding: br" & vbNewLine _
    & "Connection: keep-alive" & vbNewLine _
    & "Cookie: csrf=abcd123" & vbNewLine
    
    Dim newHost As String

   ' open internet connection
      hOpen = InternetOpen("PostmanRuntime/7.43.3", INTERNET_OPEN_TYPE_PRECONFIG, "", "", 0)
   'hOpen = InternetOpen("PostmanRuntime/7.43.3", INTERNET_OPEN_TYPE_PRECONFIG, "", "", INTERNET_FLAG_ASYNC)
   If hOpen <> 0 Then
      hConnection = InternetConnect(hOpen, PROXY, INTERNET_DEFAULT_HTTP_PORT, "", "", INTERNET_SERVICE_HTTP, 0, 0)
      If hConnection <> 0 Then
      
         hURL = HttpOpenRequest(hConnection, "POST", CGI, 0, 0, 0, INTERNET_FLAG_RELOAD, 0)
         
         'MsgBox hURL
         If hURL <> 0 Then

            If HttpSendRequest(hURL, sHeaders, Len(sHeaders), Body, Len(Body)) Then

               Dim totalData() As Byte
               ReDim totalData(0)
               
               Do
                If InternetReadFile(hURL, buffer(0), 4096, bytesRead) = 0 Then Exit Do
                If bytesRead = 0 Then Exit Do
                
                ReDim chunk(0 To bytesRead - 1)
                Dim i As Long
                For i = 0 To bytesRead - 1
                    chunk(i) = buffer(i)
                Next
                
                If firstRead Then
                    allBytes = chunk
                    firstRead = False
                Else
                    Dim newLen As Long
                    newLen = totalLen + bytesRead
                    ReDim Preserve allBytes(0 To newLen - 1)
                    For i = 0 To bytesRead - 1
                        allBytes(totalLen + i) = buffer(i)
                    Next
                End If
                
                totalLen = totalLen + bytesRead
                
               Loop
            Else
               Debug.Print (Err.LastDllError)
               'syncState.Caption = SYNC_CAPTION_PREFIX & Err.LastDllError
            End If
         Else
            Debug.Print ("505")
         End If
      Else
         Debug.Print ("506")
      End If
   Else
      Debug.Print ("507")
   End If
   
   InternetCloseHandle hURL
   InternetCloseHandle hConnection
   InternetCloseHandle hOpen

End Function

