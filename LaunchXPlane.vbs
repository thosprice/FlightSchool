Option Explicit
' =========================================================================
' FlightSchool X-Plane 9 student launcher (VBScript port)
' Runs entirely as a standard user. No elevation is requested or required.
'
' This replaces LaunchXPlane.bat. The logic is unchanged; the language is
' not, specifically to eliminate two cmd.exe parser hazards that batch
' hit in testing: (1) "exit /b" only returning from an extra CALL layer
' instead of terminating the script, and (2) "goto" misbehaving when
' jumping out of a parenthesized block to a distant label. Neither
' failure mode is possible in VBScript: WScript.Quit always terminates
' immediately regardless of call depth, and there is no goto at all.
'
' Other differences from the batch version, all deliberate improvements:
'   - Downloads use WinHttp (built into Windows) instead of curl.exe, so
'     curl's presence is no longer a prerequisite.
'   - The connectivity check is an HTTPS request, not an ICMP ping, so it
'     is not defeated by labs that block ping but allow normal browsing.
'   - Hash lookups use a Dictionary keyed on the exact manifest path, so
'     the substring-collision risk from batch's findstr-based lookup does
'     not exist here even in principle.
'   - Manifest/hash file lines are trimmed and blank/whitespace-only lines
'     are skipped uniformly, and CRLF/LF line endings are normalized
'     before splitting - this directly fixes the blank-manifest-line bug
'     from testing.
'   - Failure-archive filenames use a fixed yyyy-mm-dd_hhmmss format
'     instead of %DATE%/%TIME%, so they no longer depend on the machine's
'     regional settings.
'
' Control files on GitHub (single unified manifest):
'   manifest.txt  - one relative path per line, each beginning "X-Plane\..."
'   hashes.txt    - "RelPath:SHA256HASH" one per line
' Both files MUST begin with a line of the exact form:
'   # FLIGHTSCHOOL_VERSION=yyyy.mm.dd.vv
' A missing or malformed version line is treated as a corrupted control file.
'
' LastDownloaded\ retains the manifest.txt/hashes.txt used by the most
' recent SUCCESSFUL run, for version-drift comparison and as a fail-safe
' archive location for the log when a run fails.
' =========================================================================

Dim WshShell, FSO
Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Relaunch under cscript.exe if this was double-clicked (which defaults to
' wscript.exe, whose WScript.Echo pops a dialog box per line instead of
' writing to a console). This keeps double-click convenience while
' guaranteeing real console I/O for status lines and the failure pause.
If InStr(1, WScript.FullName, "wscript.exe", 1) > 0 Then
    WshShell.Run "cscript.exe //nologo """ & WScript.ScriptFullName & """", 1, False
    WScript.Quit 0
End If

' ============================================================
' CONFIGURATION
' ============================================================
Dim LocalAppData, ComputerName, UserName
LocalAppData = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
ComputerName = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
UserName     = WshShell.ExpandEnvironmentStrings("%USERNAME%")

Dim VStoreBase, SimPrepDir, StagingDir, LastDownloadedDir, LogFile, LogMaxLines
VStoreBase        = LocalAppData & "\VirtualStore\Program Files (x86)"
SimPrepDir        = LocalAppData & "\SimPrep"
StagingDir        = SimPrepDir & "\X-Plane"
LastDownloadedDir = SimPrepDir & "\LastDownloaded"
LogFile           = SimPrepDir & "\FlightSchool_XPlane9.log"
LogMaxLines       = 2000

' Orphan cleanup (removing VirtualStore files no longer in the manifest).
' Re-enabled for teacher-account testing now that the path-matching bug
' (below) is fixed. Set back to False if anything looks wrong.
Dim EnableOrphanCleanup
EnableOrphanCleanup = True

' Public copy of the log - readable by an instructor account without logging
' into the student's profile. Redacted (no LOCALAPPDATA/TEMP paths, which
' would otherwise expose the student's Windows username). If this location
' turns out to be locked down on the lab image, change ONLY this line.
Dim PublicLogDir, PublicLogFile
PublicLogDir  = "C:\Users\Public\Documents\FlightSchool"
PublicLogFile = PublicLogDir & "\FlightSchool_XPlane9_" & ComputerName & ".log"

Dim IsoPath, SimPath
IsoPath = "C:\Program Files (x86)\X-Plane 9\X-Plane.iso"
' alternate ISO locations for testing on other computers - comment/uncomment as needed
' IsoPath = "E:\Flight School\simulator\Computer lab\X-Plane 9 Disc 1.iso"
' IsoPath = "C:\Program Files (x86)\X-Plane 9 Disc 1.iso"
SimPath = "C:\Program Files (x86)\X-Plane 9\X-Plane.exe"
' Testing at home without X-Plane installed? Comment out the two SimPath
' check lines flagged "HOME TEST" inside CheckPrerequisites, then restore
' them before this runs at school.

Dim RepoBaseUrl
' switch the order of the following two lines if githubusercontent.com is no longer available
RepoBaseUrl = "https://cdn.jsdelivr.net/gh/thosprice/FlightSchool@main"
RepoBaseUrl = "https://raw.githubusercontent.com/thosprice/FlightSchool/refs/heads/main"

' Google Form - submitted on FAILURE ONLY. Public endpoint, no auth needed.
Dim FormUrl, FormEntryTime, FormEntryComputer, FormEntryUser, FormEntryStatus, FormEntryDetail
FormUrl           = "https://docs.google.com/forms/d/e/1FAIpQLSfC-wkY7E2_AK8NixIvYpxyxaoQ00Y8HsXygxfp4Lw-Np8_nw/formResponse"
FormEntryTime     = "entry.1262006207"
FormEntryComputer = "entry.367309525"
FormEntryUser     = "entry.11615460"
FormEntryStatus   = "entry.545270873"
FormEntryDetail   = "entry.2084914416"

Dim ManifestVer, HashesVer, HashDict, ManifestPaths

' ============================================================
' MAIN
' ============================================================
EnsureFolder SimPrepDir
EnsureFolder LastDownloadedDir

RotateLogIfNeeded

If FSO.FolderExists(StagingDir) Then
    If Not IsFolderEmpty(StagingDir) Then
        LogMessage "[WARN] Leftover staging content found from a previous run - investigate if unexpected."
    End If
End If

WriteLogBoth "========================================================="
WriteLogBoth "X-Plane 9 Initialization - " & FormatLogTimestamp(Now)
WriteLogBoth "========================================================="
LogMessage "Initializing X-Plane 9 profile..."

CheckPrerequisites
MountIso
DownloadControlFiles
ParseAndCheckVersions
LoadHashesIntoDictionary
ProcessManifest

If EnableOrphanCleanup Then
    CleanupOrphans
Else
    LogMessage "[INFO] Orphan cleanup is disabled - skipping."
End If

FinalizeSuccessfulRun

LogMessage "Booting X-Plane 9 environment..."
WshShell.Run """" & SimPath & """", 1, False

WScript.Quit 0

' =========================================================================
' SUBROUTINES
' =========================================================================

' --- Create a folder, and any missing parent folders, recursively.
Sub EnsureFolder(path)
    Dim parent
    If FSO.FolderExists(path) Then Exit Sub
    parent = FSO.GetParentFolderName(path)
    If parent <> "" And Not FSO.FolderExists(parent) Then EnsureFolder parent
    FSO.CreateFolder path
End Sub

Function IsFolderEmpty(path)
    Dim f
    Set f = FSO.GetFolder(path)
    IsFolderEmpty = (f.Files.Count = 0) And (f.SubFolders.Count = 0)
End Function

' --- Read a whole text file as an array of lines, normalizing CRLF/LF/CR
'     so manifest/hash files behave the same regardless of how they were
'     saved or which tool produced them on GitHub.
Function ReadAllLines(path)
    Dim ts, content
    Set ts = FSO.OpenTextFile(path, 1, False)
    content = ts.ReadAll
    ts.Close
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    ReadAllLines = Split(content, vbLf)
End Function

Sub RotateLogIfNeeded()
    If Not FSO.FileExists(LogFile) Then Exit Sub
    Dim lines, lineCount
    lines = ReadAllLines(LogFile)
    lineCount = UBound(lines) + 1
    If lineCount > LogMaxLines Then
        On Error Resume Next
        FSO.CopyFile LogFile, LastDownloadedDir & "\Log_Archive_Previous.log", True
        FSO.DeleteFile LogFile, True
        On Error Goto 0
    End If
End Sub

' --- Write a line to ONLY the private log file (used for the session
'     banner, which the batch version also never echoed to console).
Sub WriteLogOnly(msg)
    Dim ts
    Set ts = FSO.OpenTextFile(LogFile, 8, True)
    ts.WriteLine msg
    ts.Close
End Sub

' --- Write a line to both the private and (redacted) public log, but not
'     the console. Used for the session banner so a Google Form failure
'     report's timestamp can be matched against the public per-computer
'     log without needing the student's own profile/username.
Sub WriteLogBoth(msg)
    WriteLogOnly msg
    LogPublicRedacted msg
End Sub

' --- Write a line to the console, the private log, and a redacted copy
'     of the public log (best-effort; never fatal if that fails).
Sub LogMessage(msg)
    WScript.Echo msg
    WriteLogOnly msg
    LogPublicRedacted msg
End Sub

Sub LogPublicRedacted(msg)
    On Error Resume Next
    Dim line, tempEnv
    line = Replace(msg, LocalAppData, "LOCALAPPDATA")
    tempEnv = WshShell.ExpandEnvironmentStrings("%TEMP%")
    line = Replace(line, tempEnv, "TEMP")
    If Not FSO.FolderExists(PublicLogDir) Then EnsureFolder PublicLogDir
    Dim ts
    Set ts = FSO.OpenTextFile(PublicLogFile, 8, True)
    ts.WriteLine line
    ts.Close
    On Error Goto 0
End Sub

' --- Log a critical failure, snapshot the log, notify the Google Form,
'     show the banner, wait for a keypress, and terminate the script.
'     WScript.Quit always terminates immediately, from any call depth -
'     there is no "extra call layer" hazard here as there was in batch.
Sub HaltScript(msg)
    LogMessage "[CRITICAL] " & msg

    On Error Resume Next
    FSO.CopyFile LogFile, LastDownloadedDir & "\Failure_" & BuildFailStamp() & ".log", True
    On Error Goto 0

    SubmitFailure msg

    WScript.Echo "========================================================="
    WScript.Echo "[CRITICAL ERROR] " & msg
    WScript.Echo "Please re-run the launcher or alert your instructor."
    WScript.Echo "========================================================="
    WScript.StdOut.Write "Press Enter to continue..."
    WScript.StdIn.ReadLine

    WScript.Quit 1
End Sub

Function BuildFailStamp()
    Dim n
    n = Now
    BuildFailStamp = Year(n) & "-" & Right("0" & Month(n), 2) & "-" & Right("0" & Day(n), 2) & "_" & _
        Right("0" & Hour(n), 2) & Right("0" & Minute(n), 2) & Right("0" & Second(n), 2)
End Function

Function FormatLogTimestamp(n)
    FormatLogTimestamp = Year(n) & "-" & Right("0" & Month(n), 2) & "-" & Right("0" & Day(n), 2) & " " & _
        Right("0" & Hour(n), 2) & ":" & Right("0" & Minute(n), 2) & ":" & Right("0" & Second(n), 2)
End Function

' --- Minimal percent-encoding, sufficient for computer names, usernames,
'     and short ASCII detail messages.
Function UrlEncode(s)
    Dim i, c, code, result
    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        code = Asc(c)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or c = "-" Or c = "_" Or c = "." Or c = "~" Then
            result = result & c
        ElseIf c = " " Then
            result = result & "+"
        Else
            If code < 0 Then code = code + 65536
            result = result & "%" & Right("0" & Hex(code), 2)
        End If
    Next
    UrlEncode = result
End Function

' --- POST a failure report to the Google Form. Best-effort: a failed
'     submission (e.g. no internet) must not mask the original error.
Sub SubmitFailure(detail)
    On Error Resume Next
    Dim http, body
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    body = FormEntryTime & "=" & UrlEncode(FormatLogTimestamp(Now)) & "&" & _
           FormEntryComputer & "=" & UrlEncode(ComputerName) & "&" & _
           FormEntryUser & "=" & UrlEncode(UserName) & "&" & _
           FormEntryStatus & "=" & UrlEncode("Failed") & "&" & _
           FormEntryDetail & "=" & UrlEncode(detail)
    http.Open "POST", FormUrl, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send body
    On Error Goto 0
End Sub

Function CommandExists(exeName)
    Dim sysPath
    sysPath = WshShell.ExpandEnvironmentStrings("%WINDIR%") & "\System32\" & exeName
    CommandExists = FSO.FileExists(sysPath)
End Function

' --- HTTPS reachability check. Deliberately not an ICMP ping: some lab
'     firewalls block ping while allowing normal HTTPS traffic, which
'     made the batch version's ping-based check unreliable there.
Function CheckGithubReachable()
    On Error Resume Next
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 3000, 3000, 3000, 3000
    http.Open "HEAD", "https://raw.githubusercontent.com/", False
    http.Send
    CheckGithubReachable = (Err.Number = 0)
    On Error Goto 0
End Function

Sub CheckPrerequisites()
    Dim failCount
    failCount = 0

    If Not FSO.FolderExists(LocalAppData) Then
        failCount = failCount + 1
        LogMessage "[FAIL] Directory LOCALAPPDATA unavailable"
    End If

    If Not CommandExists("certutil.exe") Then
        failCount = failCount + 1
        LogMessage "[FAIL] CertUtil unavailable"
    End If

    If Not CheckGithubReachable() Then
        failCount = failCount + 1
        LogMessage "[FAIL] Github unavailable"
    End If

    ' HOME TEST: comment out the next 4 lines while testing without X-Plane installed
    If Not FSO.FileExists(SimPath) Then
        failCount = failCount + 1
        LogMessage "[FAIL] Can't find SimPath"
    End If

    If Not FSO.FileExists(IsoPath) Then
        failCount = failCount + 1
        LogMessage "[FAIL] Can't find IsoPath"
    End If

    If failCount > 0 Then
        HaltScript CStr(failCount) & " prerequisite checks failed"
    End If
    LogMessage "[OK] Prerequisite checks passed."
End Sub

Function IsIsoMounted()
    Dim drives, d, vol, found
    drives = Array("D", "E", "F", "G", "H")
    found = False
    For Each d In drives
        If FSO.DriveExists(d & ":") Then
            On Error Resume Next
            vol = FSO.GetDrive(d & ":").VolumeName
            On Error Goto 0
            If InStr(1, vol, "XPLANE9", 1) > 0 Then
                found = True
                Exit For
            End If
        End If
    Next
    IsIsoMounted = found
End Function

Sub MountIso()
    LogMessage "[ISO] Checking if virtual DVD is mounted..."
    If Not IsIsoMounted() Then
        LogMessage "[ISO] Virtual DVD not detected. Mounting now..."
        WshShell.Run "explorer.exe """ & IsoPath & """", 1, False
        WScript.Sleep 3000
        If Not IsIsoMounted() Then
            HaltScript "ISO mount failed"
        End If
    End If
    LogMessage "[OK] Virtual DVD confirmed mounted."
End Sub

' --- Binary-safe download via WinHttp + ADODB.Stream (text-mode writes
'     would corrupt any non-text file in the manifest).
Function DownloadFile(url, destPath)
    On Error Resume Next
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send
    If Err.Number <> 0 Or http.Status <> 200 Then
        DownloadFile = False
        Exit Function
    End If

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile destPath, 2 ' overwrite
    stream.Close

    DownloadFile = (Err.Number = 0)
    On Error Goto 0
End Function

Sub DownloadControlFiles()
    LogMessage "[REPO] Downloading control files..."
    If Not DownloadFile(RepoBaseUrl & "/manifest.txt", SimPrepDir & "\manifest.txt") Then
        HaltScript "Could not download manifest.txt"
    End If
    If Not DownloadFile(RepoBaseUrl & "/hashes.txt", SimPrepDir & "\hashes.txt") Then
        HaltScript "Could not download hashes.txt"
    End If
End Sub

Function ReadVersionHeader(path)
    ReadVersionHeader = ""
    If Not FSO.FileExists(path) Then Exit Function
    Dim lines, firstLine, prefix
    lines = ReadAllLines(path)
    If UBound(lines) < 0 Then Exit Function
    firstLine = Trim(lines(0))
    prefix = "# FLIGHTSCHOOL_VERSION="
    If Left(firstLine, Len(prefix)) = prefix Then
        ReadVersionHeader = Mid(firstLine, Len(prefix) + 1)
    End If
End Function

Sub ParseAndCheckVersions()
    ManifestVer = ReadVersionHeader(SimPrepDir & "\manifest.txt")
    If ManifestVer = "" Then HaltScript "manifest.txt missing or malformed version header"

    HashesVer = ReadVersionHeader(SimPrepDir & "\hashes.txt")
    If HashesVer = "" Then HaltScript "hashes.txt missing or malformed version header"

    If ManifestVer <> HashesVer Then
        HaltScript "manifest.txt version " & ManifestVer & " does not match hashes.txt version " & HashesVer
    End If
    LogMessage "[OK] Control file versions match: " & ManifestVer

    Dim prevVer
    If FSO.FileExists(LastDownloadedDir & "\manifest.txt") Then
        prevVer = ReadVersionHeader(LastDownloadedDir & "\manifest.txt")
        If prevVer <> "" And prevVer <> ManifestVer Then
            LogMessage "[WARN] manifest.txt version changed since last successful run: " & prevVer & " -> " & ManifestVer
        End If
    End If
    If FSO.FileExists(LastDownloadedDir & "\hashes.txt") Then
        prevVer = ReadVersionHeader(LastDownloadedDir & "\hashes.txt")
        If prevVer <> "" And prevVer <> HashesVer Then
            LogMessage "[WARN] hashes.txt version changed since last successful run: " & prevVer & " -> " & HashesVer
        End If
    End If
End Sub

' --- Load hashes.txt into a Dictionary keyed on the EXACT manifest path.
'     This removes the substring-collision risk entirely (no findstr-style
'     matching is involved at all).
Sub LoadHashesIntoDictionary()
    Set HashDict = CreateObject("Scripting.Dictionary")
    HashDict.CompareMode = 1 ' vbTextCompare
    Dim lines, i, ln, colonPos, key, val
    lines = ReadAllLines(SimPrepDir & "\hashes.txt")
    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        If ln <> "" And Left(ln, 1) <> "#" Then
            colonPos = InStr(ln, ":")
            If colonPos > 0 Then
                key = Left(ln, colonPos - 1)
                val = Mid(ln, colonPos + 1)
                If Not HashDict.Exists(key) Then HashDict.Add key, val
            End If
        End If
    Next
End Sub

Function GetFileHash(path)
    GetFileHash = ""
    On Error Resume Next
    Dim exec, output, lines
    Set exec = WshShell.Exec("certutil.exe -hashfile """ & path & """ SHA256")
    Do While exec.Status = 0
        WScript.Sleep 50
    Loop
    output = exec.StdOut.ReadAll
    On Error Goto 0

    lines = Split(Replace(output, vbCr, ""), vbLf)
    If UBound(lines) >= 1 Then
        GetFileHash = Trim(Replace(lines(1), " ", ""))
    End If
End Function

' --- Process every file in the manifest: verify existing files by hash,
'     download+verify anything missing or mismatched, halt on any
'     download or hash failure. Trimming each line here (and treating a
'     whitespace-only line the same as a blank one) is what fixes the
'     blank-manifest-line bug found during testing.
Sub ProcessManifest()
    LogMessage "[INFO] Syncing user profile data..."

    Set ManifestPaths = CreateObject("Scripting.Dictionary")
    ManifestPaths.CompareMode = 1 ' vbTextCompare

    Dim lines, i, relPath, winRelPath, targetFile, targetDir
    Dim expectedHash, needDownload, stagingFile, stagingDir, localHash, postHash, urlRelPath

    lines = ReadAllLines(SimPrepDir & "\manifest.txt")
    For i = 0 To UBound(lines)
        relPath = Trim(lines(i))
        If relPath <> "" And Left(relPath, 1) <> "#" Then
            winRelPath = Replace(relPath, "/", "\")
            If Not ManifestPaths.Exists(winRelPath) Then ManifestPaths.Add winRelPath, True

            targetFile = VStoreBase & "\" & winRelPath
            targetDir = FSO.GetParentFolderName(targetFile)
            EnsureFolder targetDir

            If Not HashDict.Exists(relPath) Then
                HaltScript "No hash entry found for " & relPath
            End If
            expectedHash = HashDict(relPath)

            needDownload = True
            If FSO.FileExists(targetFile) Then
                localHash = GetFileHash(targetFile)
                If LCase(localHash) = LCase(expectedHash) Then needDownload = False
            End If

            If needDownload Then
                stagingFile = SimPrepDir & "\" & winRelPath
                stagingDir = FSO.GetParentFolderName(stagingFile)
                EnsureFolder stagingDir

                urlRelPath = Replace(relPath, " ", "%20")
                urlRelPath = Replace(urlRelPath, "#", "%23")

                If Not DownloadFile(RepoBaseUrl & "/" & urlRelPath, stagingFile) Then
                    HaltScript "Download failed: " & relPath
                End If

                postHash = GetFileHash(stagingFile)
                If LCase(postHash) <> LCase(expectedHash) Then
                    HaltScript "Hash verification failed after download: " & relPath
                End If

                Dim copyOk
                copyOk = True
                On Error Resume Next
                FSO.CopyFile stagingFile, targetFile, True
                If Err.Number <> 0 Then copyOk = False
                On Error Goto 0
                If Not copyOk Then
                    HaltScript "Could not copy verified file into VirtualStore: " & relPath
                End If

                LogMessage "[OK] Downloaded and verified: " & relPath
            Else
                LogMessage "[OK] Verified existing: " & relPath
            End If
        End If
    Next
End Sub

Sub WalkAndCleanup(folderPath)
    Dim folder, f, sub1, relPath
    Set folder = FSO.GetFolder(folderPath)
    For Each f In folder.Files
        relPath = Mid(f.Path, Len(VStoreBase) + 2)
        If Not ManifestPaths.Exists(relPath) Then
            On Error Resume Next
            FSO.DeleteFile f.Path, True
            On Error Goto 0
            LogMessage "[CLEANUP] Removed orphaned file (not in current manifest): " & relPath
        End If
    Next
    For Each sub1 In folder.SubFolders
        WalkAndCleanup sub1.Path
    Next
End Sub

Sub CleanupOrphans()
    Dim orphanRoot
    orphanRoot = VStoreBase & "\X-Plane 9"
    If FSO.FolderExists(orphanRoot) Then WalkAndCleanup orphanRoot
End Sub

Sub FinalizeSuccessfulRun()
    On Error Resume Next
    If FSO.FolderExists(StagingDir) Then FSO.DeleteFolder StagingDir, True
    FSO.CopyFile SimPrepDir & "\manifest.txt", LastDownloadedDir & "\manifest.txt", True
    FSO.CopyFile SimPrepDir & "\hashes.txt", LastDownloadedDir & "\hashes.txt", True
    On Error Goto 0
    LogMessage "[OK] All files verified. Manifest processing complete."
End Sub
