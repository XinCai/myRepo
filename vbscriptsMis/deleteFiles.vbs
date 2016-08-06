'delete files in temp/testing folder
Set obj = CreateObject("Scripting.FileSystemObject") 'Calls the File System Object

obj.DeleteFile("C:\Game\*.txt") 'Deletes the file throught the DeleteFile function