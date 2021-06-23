FOLDER = ""	'Çopy path from Explorer terminal, save this file and open [Folder@File.bat].

if FOLDER = "" then 
	msgbox "Please edit string [FOLDER] in file [Folder@File.vbs]."
	WScript.Quit
end if

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(FOLDER & "\")
Set objSubFolders = objFolder.SubFolders
Val1 = "Folder@File.vbs"
Val2 = "Folder@File.bat"
Val3 = "Folder@File.Log"
For each objSubFolder In objSubFolders
    wscript.echo objSubFolder.name
Next

Set files = objFolder.Files
  For Each file in files
	if file.name = Val1 then 
		'Nothing
	elseif file.name = Val2 then 
		'Nothing
	elseif file.name = Val3 then 
		'Nothing
	else
		wscript.echo file.name
	end if
  Next
