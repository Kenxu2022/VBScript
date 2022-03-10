set FSO=createobject("scripting.filesystemobject")
FSO.createtextfile("D:/test.txt") 'Create a text file named "test.txt"
FSO.movefile "D:/test.txt","F:/test.txt" 'Move this file from D:\ to F:\
if FSO.fileexists("D:/test.txt") then 'verify whether this file exists in this path
	wscript.echo "File exists."
	else
		wscript.echo "File does not exists."
end if

FSO.movefile "F:/test.txt","D:/test.txt"
if FSO.fileexists("D:/test.txt") then
	wscript.echo "File exists."
	else
		wscript.echo "File does not exists."
end if

const forappending=8
set file=FSO.opentextfile("D:/test.txt",forappending,true) 'open this file as writable
file.writeline "this is a test file" 'add the following content after the last line of the file
file.close

const forreading=1
set file=FSO.opentextfile("D:/test.txt",forreading,true) 'open this file for reading
strContents=file.readall 'read the text file
msgbox "Contents:" & strcontents,vbInformation,"TEST"
file.close

wscript.echo "Test complete, file 'test.txt' will be deleted."
FSO.deletefile("D:/test.txt") 'delete this file