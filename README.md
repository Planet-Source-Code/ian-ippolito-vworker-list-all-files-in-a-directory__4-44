<div align="center">

## List all files in a directory


</div>

### Description

Uses the new file FileSystemObject in the scripting library to list all the files in the c:\inetpub\scripts\ directory with a link to them. You can modify this code to list all the files in any directory.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Intermediate
**User Rating**    |4.9 (54 globes from 11 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-list-all-files-in-a-directory__4-44/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<HTML>
<BODY>
<%
Dim objFileScripting, objFolder
dim filename, filecollection, strDirectoryPath, strUrlPath
	strDirectoryPath="c:\inetpub\scripts\"
	strUrlPath="\scripts\"
	'get file scripting object
	Set objFileScripting = CreateObject("Scripting.FileSystemObject")
	'Return folder object
	Set objFolder = objFileScripting.GetFolder("c:\inetpub\scripts\")
	'return file collection in folder
	Set filecollection = objFolder.Files
	'create the links
	For Each filename in filecollection
		Filename=right(Filename,len(Filename)-InStrRev(Filename, "\"))
		Response.Write "<A HREF=""" & strUrlPath & filename & """>" & filename & "</A><BR>"
	Next
%>
</BODY>
</HTML>
```

