Here's how I implemented Automatic Update prompts for MS Access Database Apps

Create a Github repo for your  version.json  file

Save this vba code to a bas Module in your Access database and fill in the required info.

```
' =============================================================================
'  modUpdateChecker.bas  -  GitHub-Based Update Checker for MS Access
' =============================================================================
'
'  SETUP INSTRUCTIONS (5 steps):
'  -----------------------------------------------------------------------
'  1. Import this module into your Access database (File > Import > Module)
'
'  2. In the CONFIG section below, set:
'       GITHUB_USER          - your GitHub username
'       GITHUB_REPO          - your GitHub repository name
'       GITHUB_VERSION_FILE  - path to version file in repo (default: "version.json")
'
'  3. Create a file called  version.json  in your GitHub repo with this content:
'       {
'         "version": "1.0.0",
'         "release_date": "2026-02-28",
'         "release_notes": "Initial release.",
'         "download_url": "https://github.com/YOUR-USERNAME/YOUR-REPO/releases"
'       }
'
'  4. Make sure the file is committed and pushed to GitHub.
'
'  5. Call  CheckForUpdates  from your main form's  Form_Open  event:
'       Private Sub Form_Open(Cancel As Integer)
'           CheckForUpdates         ' silent mode  - only alerts if update found
'           ' CheckForUpdates True  ' verbose mode - also tells user they are up to date
'       End Sub
'
' =============================================================================

Option Compare Database
Option Explicit

' =============================================================================
'  CONFIG  -  Edit these values for your app
' =============================================================================

Private Const GITHUB_USER          As String = "YOUR_GITHUB_USERNAME"             ' <-- YOUR GitHub username
Private Const GITHUB_REPO          As String = "YOUR_REPO_NAME"       ' <-- YOUR repo name
Private Const GITHUB_VERSION_FILE  As String = "version.json"     ' path inside repo
Private Const CHECK_TIMEOUT_SEC    As Long = 8                    ' HTTP timeout in seconds

' =============================================================================
'  GET CURRENT VERSION
' =============================================================================

Private Function GetCurrentVersion() As String
    GetCurrentVersion = CStr(Nz(DLookup("[Version]", "4DBSettings", "[ID] = 1"), "0.0.0"))         ' <-- Current version stored in a table
End Function

' =============================================================================
'  GET APP DOWNLOAD URL
' =============================================================================

Private Function AppDownloadURL() As String
    ' Override download URL (leave "" to use value in version.json on Github)
    ' direct download URL should end with "?download=1"
    AppDownloadURL = DLookup("[AppDownloadFolder]", "4DBSettings", "[ID] = 1")         ' <-- App Download URL stored in a table
End Function

' =============================================================================
'  PUBLIC ENTRY POINT
' =============================================================================

' CheckForUpdates
'   silent (default = True):
'     True  - only show a dialog when an update IS available (good for startup)
'     False - also confirm when the user is already on the latest version
'
Public Sub CheckForUpdates(Optional Silent As Boolean = True)

    Dim sJson       As String
    Dim sLatest     As String
    Dim sNotes      As String
    Dim sDate       As String
    Dim sDownload   As String
    Dim sMsg        As String
    Dim nResult     As Integer

    ' --- fetch version.json from GitHub raw content ---
    sJson = FetchVersionJson()
    If sJson = "" Then
        If Not Silent Then
            MsgBox "Could not reach GitHub to check for updates." & vbCrLf & _
                   "Please check your internet connection and try again.", _
                   vbInformation, "Update Check"
        End If
        Exit Sub
    End If

    ' --- parse fields ---
    sLatest = JsonField(sJson, "version")
    sDate = JsonField(sJson, "release_date")
    sNotes = JsonField(sJson, "release_notes")
    sDownload = JsonField(sJson, "download_url")

    If sLatest = "" Then
        If Not Silent Then
            MsgBox "The version file on GitHub could not be read." & vbCrLf & _
                   "Please check the version.json format.", _
                   vbExclamation, "Update Check"
        End If
        Exit Sub
    End If

    ' --- compare ---
    If IsNewerVersion(sLatest, GetCurrentVersion()) Then

        ' Build update message
        sMsg = "A new version is available!" & vbCrLf & vbCrLf & _
               "  Installed :  " & GetCurrentVersion() & vbCrLf & _
               "  Available :  " & sLatest

        If sDate <> "" Then
            sMsg = sMsg & "  (" & sDate & ")"
        End If

        If sNotes <> "" Then
            sMsg = sMsg & vbCrLf & vbCrLf & _
                   "What's new:" & vbCrLf & "  " & sNotes
        End If

        sMsg = sMsg & vbCrLf & vbCrLf & "Would you like to go to the download page?"

        nResult = MsgBox(sMsg, vbYesNo + vbInformation, "Update Available")

        If nResult = vbYes Then
            ' Prefer override URL, then version.json URL, then default
            Dim sUrl As String
            If AppDownloadURL() <> "" Then
                sUrl = AppDownloadURL()
            ElseIf sDownload <> "" Then
                sUrl = sDownload
            Else
                sUrl = "https://github.com/" & GITHUB_USER & "/" & GITHUB_REPO & "/releases"
            End If
            Application.FollowHyperlink sUrl
        End If

    Else
        If Not Silent Then
            MsgBox "You are running the latest version (" & GetCurrentVersion() & ").", _
                   vbInformation, "No Updates Available"
        End If
    End If

End Sub

' =============================================================================
'  PRIVATE HELPERS
' =============================================================================

' FetchVersionJson
'   Downloads the raw version.json from GitHub and returns its text.
'   Returns "" on any error so callers can handle gracefully.
'
Private Function FetchVersionJson() As String

    Dim sUrl    As String
    Dim oHttp   As Object
    Dim sResult As String

    ' GitHub raw content URL - using refs/heads/main for reliability
    sUrl = "https://raw.githubusercontent.com/" & _
           GITHUB_USER & "/" & GITHUB_REPO & "/refs/heads/main/" & GITHUB_VERSION_FILE

    On Error GoTo Fail

    ' ServerXMLHTTP handles HTTPS/TLS correctly; XMLHTTP can fail on modern GitHub URLs
    Set oHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    oHttp.Open "GET", sUrl, False
    oHttp.setTimeouts CHECK_TIMEOUT_SEC * 1000, _
                       CHECK_TIMEOUT_SEC * 1000, _
                       CHECK_TIMEOUT_SEC * 1000, _
                       CHECK_TIMEOUT_SEC * 1000
    oHttp.Send

    If oHttp.Status = 200 Then
        sResult = oHttp.responseText
    End If

    Set oHttp = Nothing
    FetchVersionJson = sResult
    Exit Function

Fail:
    FetchVersionJson = ""
    On Error Resume Next
    Set oHttp = Nothing

End Function

' JsonField
'   Minimal JSON string parser - extracts a top-level string field value.
'   Handles:  "key": "value"  (with or without spaces around the colon)
'   Does NOT handle nested objects or arrays (not needed for version.json).
'
Private Function JsonField(sJson As String, sKey As String) As String

    Dim nKey    As Long
    Dim nColon  As Long
    Dim nOpen   As Long
    Dim nClose  As Long
    Dim sValue  As String

    ' Find  "key"
    nKey = InStr(sJson, """" & sKey & """")
    If nKey = 0 Then Exit Function

    ' Find the colon after the key
    nColon = InStr(nKey + Len(sKey) + 2, sJson, ":")
    If nColon = 0 Then Exit Function

    ' Find opening quote of value
    nOpen = InStr(nColon + 1, sJson, """")
    If nOpen = 0 Then Exit Function

    ' Find closing quote (skip escaped quotes \")
    nClose = nOpen + 1
    Do While nClose <= Len(sJson)
        If Mid(sJson, nClose, 1) = """" And Mid(sJson, nClose - 1, 1) <> "\" Then
            Exit Do
        End If
        nClose = nClose + 1
    Loop

    If nClose > Len(sJson) Then Exit Function

    sValue = Mid(sJson, nOpen + 1, nClose - nOpen - 1)

    ' Unescape basic JSON escapes
    sValue = Replace(sValue, "\""", """")
    sValue = Replace(sValue, "\\", "\")
    sValue = Replace(sValue, "\/", "/")
    sValue = Replace(sValue, "\n", vbCrLf)
    sValue = Replace(sValue, "\r", "")
    sValue = Replace(sValue, "\t", vbTab)

    JsonField = Trim(sValue)

End Function

' IsNewerVersion
'   Returns True if sRemote is a higher version than sLocal.
'   Supports semantic versioning: MAJOR.MINOR.PATCH  (e.g. "2.1.0" vs "2.0.9")
'   Also handles single ("1") and two-part ("1.2") version strings.
'
Private Function IsNewerVersion(sRemote As String, sLocal As String) As Boolean

    Dim aParts(1 To 2, 1 To 3) As Long   ' (remote/local, major/minor/patch)
    Dim aRemote()   As String
    Dim aLocal()    As String
    Dim I           As Integer

    aRemote = Split(sRemote, ".")
    aLocal = Split(sLocal, ".")

    ' Fill remote parts (up to 3)
    For I = 0 To UBound(aRemote)
        If I > 2 Then Exit For
        aParts(1, I + 1) = Val(aRemote(I))
    Next I

    ' Fill local parts (up to 3)
    For I = 0 To UBound(aLocal)
        If I > 2 Then Exit For
        aParts(2, I + 1) = Val(aLocal(I))
    Next I

    ' Compare major, then minor, then patch
    For I = 1 To 3
        If aParts(1, I) > aParts(2, I) Then
            IsNewerVersion = True
            Exit Function
        ElseIf aParts(1, I) < aParts(2, I) Then
            IsNewerVersion = False
            Exit Function
        End If
    Next I

    ' Equal version
    IsNewerVersion = False

End Function

' =============================================================================
'  OPTIONAL: Manual check triggered from a menu button or Help > Check Updates
' =============================================================================

' Use this from a button/menu item - always shows a result (non-silent)
Public Sub ManualCheckForUpdates()
    CheckForUpdates Silent:=False
End Sub
```
