<!--#include file="inc/inc_syssite.asp"-->
<%
dim formsize,mydata,freesize,maxsize,FilePath,my_name,filename,upload_dir
Dim rs,fileID
Response.buffer=true
formsize=Request.totalbytes
mydata=Request.binaryread(formsize)
if not oblog.checkuserlogined() then
	Response.Write("登录后才能保存大头贴")
	Response.End()
end If
upload_dir=oblog.CacheConfig(56)
maxsize=oblog.l_Group(24,0)
freesize=Int(maxsize-oblog.l_uUpUsed/1024)
if formsize/1024>freesize And maxsize>0 Then
	Response.Write("空间已满，请清理上传文件。")
	Response.end
end if
if upload_dir<>"" then
	FilePath=upload_dir
else
	FilePath = oblog.l_udir&"/"&oblog.l_ufolder&"/upload"
end if

FilePath=CreatePath(FilePath)
filename=FormatName("jpg")
my_name= FilePath&filename
Call SaveStream(my_name,mydata)
oblog.execute("update oblog_user set user_upfiles_size=user_upfiles_size+"&formsize&" where userid="&oblog.l_uid)
oblog.execute("Insert into oblog_upfile (userid,file_name,file_path,file_ext,file_size) values ("&oblog.l_uid&",'"&filename&"','"&my_name&"','jpg',"&formsize&")")
Set rs = oblog.Execute ("select FileID FROM oblog_upfile WHERE file_name = '"&filename&"' ")
fileID = rs(0)
rs.Close
Set rs = Nothing
oblog.execute("Insert INTO oblog_album (userid,photo_Name,photo_path,sysclassid,userclassid,isBigHead,fileID) values ("&oblog.l_uid&",'"&filename&"','"&my_name&"',0,0,1,"&fileID&")")
Response.Write my_name

Sub SaveStream(paR_strFile, paR_streamContent)
	Dim objStream
	Set objStream =Server.CreateObject(oblog.CacheCompont(2))
		with objStream
			.Type =1
			.Open
			.Write paR_streamContent
			.SaveToFile Server.Mappath(paR_strFile), 2
			.Close()
		End with
	Set objStream =Nothing
End Sub

'检查上传目录，若无目录则自动建立
Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	if upload_dir<>"" then
		uploadpath = year(Date) & "-" & month(Date)
	else
		uploadpath=""
	end if
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	On Error Resume Next
	Set objFSO = Server.CreateObject(oblog.CacheCompont(1))
		If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
		End If
		If Err.Number = 0 and upload_dir<>"" Then
			CreatePath = PathValue & uploadpath & "/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function

Function FormatName(Byval FileExt)
	Dim RanNum,TempStr
	Randomize
	RanNum = Int(900000*rnd)+100000
	'TempStr = Year(now) & Month(now) & Day(now) & RanNum & "." & FileExt
	TempStr = Month(now) & Day(now) & RanNum & "." & FileExt
	FormatName = TempStr
End Function
%>
