Set objHTTP = CreateObject("MSXML2.XMLHTTP")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Read URLs from a text file
urlFile = "urls.txt"
Dim urlList
urlList = ""

If Not objFSO.FileExists(urlFile) Then
    WScript.Echo "URL file not found: " & urlFile
    ' Prompt user to enter a URL
    urlList = InputBox("Please enter a URL to download files:")
Else
    Set objFile = objFSO.OpenTextFile(urlFile, 1)
    If objFile.AtEndOfStream Then
        WScript.Echo "URL file is empty. Please enter a URL to download files:"
        urlList = InputBox("Please enter a URL to download files:")
    Else
        ' Read URLs from the file
        Do While Not objFile.AtEndOfStream
            urlList = urlList & Trim(objFile.ReadLine) & vbCrLf
        Loop
        objFile.Close
    End If
End If

' Split the URLs into an array
Dim urls
urls = Split(urlList, vbCrLf)

' Create downloads folder
If Not objFSO.FolderExists("downloads") Then
    objFSO.CreateFolder("downloads")
End If

' Process each URL in the array
For Each URL In urls
    URL = Trim(URL)
    
    If URL <> "" Then
        ' Confirm with the user before proceeding
        If MsgBox("You are about to download files from: " & URL & ". Do you want to continue?", vbYesNo) = vbNo Then
            WScript.Echo "Download canceled for: " & URL
        Else
            ' Download the webpage
            objHTTP.Open "GET", URL, False
            objHTTP.Send
            HTMLContent = objHTTP.ResponseText

            ' Extract and download files
            Set objRegEx = CreateObject("VBScript.RegExp")
            objRegEx.Pattern = "href=""([^""]*\.(xlsx|xls|pdf|docx|doc|zip|csv))"
            objRegEx.IgnoreCase = True
            objRegEx.Global = True

            Set objMatches = objRegEx.Execute(HTMLContent)
            If objMatches.Count = 0 Then
                WScript.Echo "No downloadable files found on: " & URL
            Else
                For Each objMatch In objMatches
                    FileURL = objMatch.SubMatches(0)
                    
                    ' Check if the FileURL starts with "/"
                    If Left(FileURL, 1) = "/" Then
                        FileURL = "https://www.epa.gov" & FileURL
                    End If
                    
                    WScript.Echo "Downloading: " & FileURL
                    
                    ' Check if the FileURL is valid
                    If InStr(FileURL, "http") = 0 Then
                        WScript.Echo "Invalid file URL: " & FileURL
                    Else
                        objHTTP.Open "GET", FileURL, False
                        On Error Resume Next
                        objHTTP.Send
                        If Err.Number <> 0 Then
                            WScript.Echo "Error downloading: " & FileURL & " (Error: " & Err.Description & ")"
                            Err.Clear
                        Else
                            ' Check if the request was successful
                            If objHTTP.Status = 200 Then
                                ' Save to file using binary stream
                                FileName = Right(FileURL, Len(FileURL) - InStrRev(FileURL, "/"))
                                
                                ' Ensure the filename is less than 255 characters
                                If Len(FileName) >= 255 Then
                                    ' Truncate the filename to 250 characters and add an ellipsis
                                    FileName = Left(FileName, 250) & "..."
                                End If
                                
                                ' Ensure the full path is less than 255 characters
                                Dim fullPath
                                fullPath = "downloads\" & FileName
                                If Len(fullPath) >= 255 Then
                                    ' Truncate the filename to ensure the full path is valid
                                    FileName = Left(FileName, 255 - Len("downloads\") - 1) & "..."
                                End If
                                
                                Set objStream = CreateObject("ADODB.Stream")
                                objStream.Type = 1 ' Binary
                                objStream.Open
                                objStream.Write objHTTP.ResponseBody
                                objStream.SaveToFile "downloads\" & FileName, 2 ' Overwrite if exists
                                objStream.Close
                            Else
                                WScript.Echo "Failed to download: " & FileURL & " (Status: " & objHTTP.Status & ")"
                            End If
                        End If
                        On Error GoTo 0
                    End If
                Next
                WScript.Echo "Downloads complete for: " & URL
            End If
        End If
    End If
Next

WScript.Echo "All downloads complete!"
