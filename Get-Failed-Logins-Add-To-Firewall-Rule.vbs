' VBScript to extract IP addresses from Event ID 4625 in the Security log and add to Windows Firewall rule

Option Explicit

Dim objWMIService, colLoggedEvents, objEvent, objRegEx, Matches, Match
Dim strComputer, strQuery, IPAddress, objFSO, objFile
Dim dictIPs, i, insertionString, existingIPs, newIPs, cmd, objShell

' Set up the regular expression to capture IP addresses
Set objRegEx = New RegExp
objRegEx.IgnoreCase = True
objRegEx.Global = True
objRegEx.Pattern = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"

' Initialize dictionary to store unique IPs
Set dictIPs = CreateObject("Scripting.Dictionary")

' Define file system object to write to a text file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("C:\unique_ip_addresses.txt", True)

' WMI service for security event log
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

' Query for Event ID 4625 in Security Log
strQuery = "Select * from Win32_NTLogEvent Where Logfile = 'Security' And EventCode = '4625'"
Set colLoggedEvents = objWMIService.ExecQuery(strQuery)

' Loop through all events
For Each objEvent In colLoggedEvents
    ' Loop through each insertion string in the event log entry
    For i = 0 To UBound(objEvent.InsertionStrings)
        insertionString = objEvent.InsertionStrings(i)
        
        ' Search the insertion string for any IP addresses
        Set Matches = objRegEx.Execute(insertionString)
        For Each Match In Matches
            IPAddress = Match.Value
            ' Add the IP address to the dictionary if not already present
            If Not dictIPs.Exists(IPAddress) Then
                dictIPs.Add IPAddress, True
            End If
        Next
    Next
Next

' Write unique IP addresses to the text file
For Each IPAddress In dictIPs.Keys
    objFile.WriteLine IPAddress
Next

' Clean up the file operations
objFile.Close
Set objFile = Nothing
Set objFSO = Nothing

' Get current IPs from the firewall rule "01-Server Login Attempts"
Set objShell = CreateObject("WScript.Shell")
Set objFile = objFSO.CreateTextFile("C:\firewall_rule_ips.txt", True)

' Run netsh command to get the current rule settings
cmd = "netsh advfirewall firewall show rule name=""01-Server Login Attempts"""
Set objExec = objShell.Exec(cmd)

Do While Not objExec.StdOut.AtEndOfStream
    objFile.WriteLine objExec.StdOut.ReadLine()
Loop

objFile.Close

' Read the current IP addresses from the firewall rule
existingIPs = ""
Set objFile = objFSO.OpenTextFile("C:\firewall_rule_ips.txt", 1)

Do While Not objFile.AtEndOfStream
    Dim line
    line = objFile.ReadLine()
    If InStr(line, "RemoteIP") > 0 Then
        existingIPs = Trim(Split(line, ":")(1))
    End If
Loop

objFile.Close
Set objFile = Nothing

' Combine existing IPs with the new ones
newIPs = existingIPs
For Each IPAddress In dictIPs.Keys
    If InStr(newIPs, IPAddress) = 0 Then
        If Len(newIPs) > 0 Then
            newIPs = newIPs & ","
        End If
        newIPs = newIPs & IPAddress
    End If
Next

' Apply the new list of IPs to the firewall rule
cmd = "netsh advfirewall firewall set rule name=""01-Server Login Attempts"" new remoteip=" & newIPs
objShell.Run cmd, 0, True

' Clean up
Set objShell = Nothing
Set dictIPs = Nothing
Set objRegEx = Nothing
Set colLoggedEvents = Nothing
Set objWMIService = Nothing

WScript.Echo "IP addresses extracted, saved to C:\unique_ip_addresses.txt, and added to the Windows Firewall rule '01-Server Login Attempts'."
