<!-- #include file="inc/inc_syssite.asp" -->
<!-- #include File="inc/class_upfile.asp" -->
<%
'------------------------------------------------
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
if not oblog.checkuserlogined() And Not oblog.CheckAdmin(0) then
	Response.Write("��¼������ϴ��ļ�")
	Response.End()
end If
Dim tMode,rs,sysclass,subject,upload_dir,ShowName,TeamID,WriteData,isencomment
Dim freesize,onesize,maxsize,enupload,upfiletype,re,isphoto
Dim fileID,ishide,photo_Readme,photo_title
Dim upwatermark
upload_dir=oblog.CacheConfig(56)
tMode=Request("tMode")
re=Request.QueryString("re")
isphoto=CInt(Request.QueryString("isphoto"))
sysclass=Request.QueryString("sysclass")
subject=Request.QueryString("subject")
ShowName=ProtectSQL(Request.QueryString("ShowName"))
photo_Readme=ProtectSQL(Request.QueryString("photo_readme"))
photo_title=ProtectSQL(Request.QueryString("photo_title"))
TeamID=Request("TeamID")
ishide=Request("ishide")
isencomment=Request("isencomment")

if sysclass<>"" then sysclass=CLng(sysclass) else sysclass=0
if subject<>"" then subject=CLng(subject) else subject=0
If TeamID <> "" Then TeamID = CLng(TeamID) Else TeamID = 0
If ishide <> "" Then ishide = CLng(ishide) Else ishide = 0
If isencomment <> "" Then isencomment = CLng(isencomment) Else isencomment = 1
If photo_title <> "" Then photo_title = Left (photo_title,30)
If photo_Readme <> "" Then photo_Readme = Left (photo_Readme,500)
If ShowName <> "" Then ShowName = Right(ShowName,30)
If tMode = "8" Then
	If GroupManageID = False Then
		Response.Write "<font color=red>����ͨ���󣬷����ϴ��Զ���ͼƬ</font>"
		Response.End
	End If
End If
'�ϴ�Ⱥ�����ͷ��
If tMode = "9" Or tMode = "8" Then
	WriteData = False
Else
	WriteData = True
End If
If oblog.CheckAdmin(0) And Not oblog.checkuserlogined() Then
	enupload = 1
Else
	If oblog.l_Group(24,0)=-1 Then
		enupload=0
	Else
		enupload=1
	End If
	upfiletype=oblog.l_Group(22,0)
	onesize=oblog.l_Group(23,0)
	maxsize=oblog.l_Group(24,0)
	upwatermark=oblog.l_Group(25,0)
End if
If tMode = 2 Then
	If Request("t") = 1 Then
		If photo_title = ""  Then
			oblog.adderrstr ("������������Ҫ��д��Ƭ���⣡")
		End If
		If oblog.chk_badword(photo_title) > 0 Then
			oblog.adderrstr ("������Ƭ���⺬��ϵͳ��������ַ���")
		End If
		If oblog.chk_badword(photo_Readme) > 0 Then
			oblog.adderrstr ("������Ƭ���ܺ���ϵͳ��������ַ���")
		End if
	End if
	upfiletype = "gif|jpg|png"
	If TeamID > 0 Then
		If Not CheckQQMember Then
			oblog.ShowMsg ("���󣺷Ǳ�" &oblog.CacheConfig(69)& "��Ա��Ȩ������Ƭ��"),"back"
		End If
	End if
End If
If oblog.errstr<>"" Then
	oblog.showusererr
	Response.End
End if
if enupload=0 then
	Response.Write("��ǰϵͳ���ò������ϴ��ļ�")
	Response.End()
end if
'maxsize�����ƴ�С,�����м��
If maxsize<>0 Then
	freesize=Int(maxsize*1024-oblog.l_uUpUsed)
	if freesize<=0 then
		Response.Write("<ul style='margin:0px;text-align: left;width:100%;'> �ϴ��ռ��������������ϴ��ļ�,�������ϴ��ĵ�</ul></body></html>")
		Response.End()
	end If
Else
	freesize = onesize
End If%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ļ��ϴ�</title>
<%If tMode="2" Then%>
<link href="oBlogStyle/UserAdmin/7/style.css" rel="stylesheet" type="text/css" />
<%Else%>
<link href="oBlogStyle/upload.css" rel="stylesheet" type="text/css" />
<%End if%>
</head>
<body>
<%
If Request("t")="1" Then
	Upfile_Main()
Else
	if tMode=1 then
		Main_photo ()
	else
		Main()
	end if
End If

Sub Main()

	Dim PostRanNum
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form name="myform" method="post" action="upload.asp?t=1&tMode=<%=tMode%>&re=<%=re%>&isphoto=<%=isphoto%>&TeamID=<%=TeamID%>" enctype="multipart/form-data">
					<ul id="Upload">
					<%If WriteData Then%>
					<%If tMode="2" Then%>
						<li class="l1">
					<%End if%>
					<%End if%>
							<input type="hidden" name="UploadCode" value="<%=PostRanNum%>">
							<input type="hidden" name="act" value="upload">
<%If WriteData Then%>
<%If tMode="2" Then%>
							<span>����</span>
							<input type="text" name = "photo_title" id="photo_title" size="40" maxlength="30" />

							<select name="photoclass" id="photoclass">
								<%=oblog.show_class("log",0,1)%>
							</select>
<%If teamid = 0 Then %>
							<select name="subjectid" id="subjectid">
								<option value="0">�ҵķ���</option>
								<%
								Set rs = oblog.Execute("select subjectid,subjectname from oblog_subject where userid=" & oblog.l_uid & " And subjectType=1")
								While Not rs.EOF
									Response.Write ("<option value=" & rs(0) & " >" & oblog.filt_html(rs(1)) & "</option>")
									rs.movenext
								Wend
								%>
							</select>
						</li>
<%Else
if not oblog.checkuserlogined() Then Response.Clear:Response.Write "���ȵ�¼":Response.End
%>
<%End if%>
<%End if%>
<%End If%>
						<li class="l2">
							<span>�ϴ�</span>
							<input type="file" name="uploadfile" id="uploadfile">
							<input type="hidden" name="fname">
							<%If tMode<>"2" Then%>
								<input type="button" name="Ok" value="�ϴ�" onclick="return ReSubmit(this.form,this.form.uploadfile.value);"  >
							<%End if%>
							<br /><span></span>
							&nbsp;<%If WriteData Then
								Response.Write "ʣ��ռ䣺"
								If maxsize=0 Then
									Response.Write "<font class=""red"">������</font>"
								Else
									Response.Write "<font class=""blue"">"&oblog.showsize(freesize)&"</font>"
								End If
								%> �����ļ���<font class="blue"><%=oblog.showsize(onesize*1024)%> </font>
								&nbsp;�����ϴ��ļ���ʽ��<font class="blue"><%=upfiletype%></font>  ������������:100�������ַ�.
							<%end if%>
						</li>
<%If WriteData Then%>
<%If tMode="2" Then%>
						<li class="l3">
							<span>˵��</span>
							<textarea name="photo_readme" id = "photo_readme"cols="40" rows="5"></textarea>
						</li>
						<li class="l4">
							<span></span><label><input type="checkbox" name="ishide" id="ishide" />������Ƭ</label>
							<label><input type="checkbox" name="isencomment" id="isencomment" />����������</label>
						</li>

						<li class="l5"><span></span><input type="button" name="Ok" value="������Ƭ" onclick="return ReSubmit(this.form,this.form.uploadfile.value);" ></li>
					</ul>
<%End If%>
<%End If%>
				</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>

<script language="javascript">
function ReSubmit(form, file1) {
	if (!file1) {
		alert('��ѡ����Ҫ�ϴ����ļ�');
		return false;
		}
		if (!file1||file1.indexOf(":\\")==-1) {
		alert('����ȷѡ����Ҫ�ϴ����ļ�');
		return false;
		} 
	while (file1.indexOf("\\") != -1)
	file1 = file1.slice(file1.indexOf("\\") + 1);
	//ext = file1.slice(file1.indexOf(".")).toLowerCase();
	if(form.action.indexOf("&ShowName=")<=0)
	form.action=form.action+"&ShowName=" + file1;
	<%If tMode = 2 Then %>
	var photoclass = document.getElementById("photoclass").value;
	var isencomment;
	if (document.getElementById("isencomment").checked ==true){
		isencomment = 0;
	}
	else {
		isencomment = 1;
	}
	<%if teamID = 0 then%>
	var subjectid = document.getElementById("subjectid").value;
	var ishide;
	if (document.getElementById("ishide").checked ==true){
		ishide = 1;
	}
	else {
		ishide = 0;
	}
	<%end if%>
	var photo_readme = document.getElementById("photo_readme").value;
	var photo_title = document.getElementById("photo_title").value;
		if (photo_title==''){
			alert('��������Ҫ��д��Ƭ����');
			document.getElementById("photo_title").focus();
			return false;
		}

	form.action=form.action+"&sysclass=" + photoclass  +"&photo_readme=" + escape(photo_readme)+"&photo_title=" + escape(photo_title)+"&isencomment="+ isencomment;
	<%if teamID = 0 then %>
	form.action=form.action+"&ishide=" + ishide+"&subject=" + subjectid;
	<%end if%>
	<%end if%>
//	alert(form.action);
//	return false;
	form.submit();
}
</script>
<%
End Sub

sub main_photo()

	Dim PostRanNum,subjectid

	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
%>
	<ul style="margin:0px;text-align: left;width:100%;">
	ʣ��ռ䣺<%
	If maxsize=0 Then
		Response.Write "������"
	Else
		Response.Write oblog.showsize(freesize)
	End If
	%> �����ļ���<%=oblog.showsize(onesize)%>
     <form name="myform" method="post" action="upload.asp?t=1&tMode=<%=tMode%>&re=<%=re%>&isphoto=<%=isphoto%>" enctype="multipart/form-data">
	<INPUT TYPE="hidden" NAME="UploadCode" value="<%=PostRanNum%>">
	<input type="hidden" name="act" value="upload">
	�ļ���<input type="file" name="uploadfile" style="width:180px">
	<span id="mup1"></span>
	<span id="mup2"></span>
	<span id="mup3"></span>
	<span id="mup4"></span>
	<span id="mup5"></span>

	<input type="hidden" name="fname">
	<br/><br />
	<input type="submit" name="Ok" value="�ϴ���Ƭ" >
	 <input type="button" value="�����ϴ�����" onClick="addf();">
     </form>
	</ul>
</body>
<script language="javascript">
var i=0;
function addf(){
	i=i+1;
	if (i<5){
	document.all["mup"+i].innerHTML='<br />�ļ���<input type="file" name="uploadfile'+i+'" style="width:200px"> <input type="button" value="ɾ��" onclick=delm("'+i+'");>';
	}else{
	i=i-1
	alert("��������ϴ�����!")
	}

}
function delm(m){
	document.all["mup"+m].innerHTML='';
	i=i-1;
}
</script>
</html>
<%
end sub

Sub Upfile_Main()
%>
<ul style="margin:0px;text-align: left;width:100%;">
<%
UploadFile
%>
</ul>
</body>
</html>
<%
End Sub

Sub UploadFile()
	'If Not oblog.ChkPost Then
	'	Exit Sub
	'End If
	Server.ScriptTimeOut=9999999
'	'-----------------------------------------------------------------------------
	Dim Upload,FilePath,FormName,File,F_FileName,F_Viewname
	dim DrawInfo
	upfiletype=Replace(upfiletype,"|",",")
	if freesize<=onesize then onesize=freesize
	if onesize<0 then onesize=0
	'����ͷ��Ⱥ��ͼƬֻ����ͼƬ��ʽ�ļ�,��СΪ200k
	If Not WriteData Then
		onesize = 200
		upfiletype = "gif,jpg,png"
	End if
	if upload_dir<>"" then
		FilePath=upload_dir
	else
		FilePath = oblog.l_udir&"/"&oblog.l_ufolder&"/upload"
	end If
	If tMode = "9" Then FilePath = FilePath & "/UploadFace"
	If tMode = "8" Then FilePath = FilePath & "/UploadGroup"
	FilePath=CreatePath(FilePath)
	If oblog.CacheCompont(12)="1" Then
		DrawInfo = oblog.CacheCompont(13)
	ElseIf oblog.CacheCompont(12)="2" Then
		DrawInfo = oblog.CacheCompont(18)
	Else
		DrawInfo = ""
	End If
	If DrawInfo = "0" Then
		DrawInfo = ""
		oblog.CacheCompont(12) = 0
	End If
	Set Upload = New UpFile_Cls
		if isphoto=1 then
			Upload.UploadType		= 0										'�����ϴ��������
		else
			Upload.UploadType		= Cint(oblog.CacheCompont(11))			'�����ϴ��������
		end if
		Upload.UploadPath			= FilePath								'�����ϴ�·��
		Upload.MaxSize				= Int(onesize)							'��λ KB
		Upload.InceptMaxFile		= 8										'ÿ���ϴ��ļ���������
		Upload.InceptFileType		= upfiletype							'�����ϴ��ļ�����
		Upload.RName				= ""
		Upload.ChkSessionName		= "UploadCode"
		if CLng(oblog.CacheCompont(12))=1 or CLng(oblog.CacheCompont(12))=2 then
			Upload.PreviewType		= 1										'����Ԥ��ͼƬ�������
		else
			Upload.PreviewType		= 999
		end if
		Upload.PreviewImageWidth	= 130									'����Ԥ��ͼƬ���
		Upload.PreviewImageHeight	= 100									'����Ԥ��ͼƬ�߶�

		Upload.DrawImageWidth		= oblog.CacheCompont(22)				'����ˮӡͼƬ������������
		Upload.DrawImageHeight		= oblog.CacheCompont(21)				'����ˮӡͼƬ����������߶�
		Upload.DrawGraph			= oblog.CacheCompont(19)				'����ˮӡ͸����
		Upload.DrawFontColor		= oblog.CacheCompont(15)				'����ˮӡ������ɫ
		Upload.DrawFontFamily		= oblog.CacheCompont(16)				'����ˮӡ���������ʽ
		Upload.DrawFontSize			= oblog.CacheCompont(17)				'����ˮӡ���������С
		Upload.DrawFontBold			= oblog.CacheCompont(17)				'����ˮӡ�����Ƿ����
		Upload.DrawInfo				=  DrawInfo								'����ˮӡ������Ϣ��ͼƬ��Ϣ
		If upwatermark=0 Then
		Upload.DrawType				= 0
		Else
		Upload.DrawType				= oblog.CacheCompont(12)				'0=������ˮӡ ��1=����ˮӡ���֣�2=����ˮӡͼƬ
		End If
		Upload.DrawXYType			= oblog.CacheCompont(23)				'"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
		Upload.DrawSizeType			= 1										'"0"=�̶���С��"1"=�ȱ�����С
		If oblog.CacheCompont(21)<>"" or oblog.CacheCompont(20)<>"0" Then
			Upload.TransitionColor	= oblog.CacheCompont(20)				'͸������ɫ����
		End If
		If tMode = "9" Then
			Upload.FileNameByID = oblog.l_uid
		ElseIf tMode = "8" Then
			Upload.FileNameByID = TeamID
		End if
		'ִ���ϴ�
		Upload.SaveUpFile
		If Upload.ErrCodes<>0 Then
			oblog.ShowMsg "����"& Upload.Description  ,"upload.asp?re="&re&"&isphoto="&isphoto&"&tMode="& tMode &"&TeamID="&TeamID
			Exit Sub
		End If
		If Upload.Count > 0 Then
			For Each FormName In Upload.UploadFiles
				Set File = Upload.UploadFiles(FormName)
				F_FileName = FilePath & File.FileName
				'����Ԥ����ˮӡͼƬ
				If WriteData Then
					If Upload.PreviewType<>999 and File.FileType=1 then
							F_Viewname =  FilePath&"pre" & Replace(File.FileName,File.FileExt,"") & "jpg"
							'����Ԥ��ͼƬ:Call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��,ԭ�ļ���׺)
							Upload.CreateView F_FileName,F_Viewname,File.FileExt
					End If
				End if
				'д���ݿ�������˴�
				If WriteData Then
					oblog.execute("Insert into oblog_upfile (userid,file_name,file_path,file_ext,file_size,file_ShowName,isphoto,FileType) values ("&oblog.l_uid&",'"&File.FileName&"','"&F_FileName&"','"&File.FileExt&"',"&File.FileSize&",'"&ShowName&"',"&isphoto&","&file.filetype&")")
					Set rs = oblog.Execute ("select FileID FROM oblog_upfile WHERE file_name = '"&File.FileName&"' ")
					fileID = rs(0)
					rs.Close
					Set rs = Nothing
					If isphoto = 1 Then
						Dim rsS
						Set rsS = oblog.Execute ("SELECT ishide FROM oblog_subject WHERE subjectid = "&subject)
						If Not rsS.Eof Then
							If rsS(0) = 1 Then ishide = 1
						End If
						Set rsS = Nothing
						oblog.execute("Insert into oblog_album (userid,photo_Name,photo_path,sysclassid,userclassid,fileID,photo_Readme,ishide,photo_title,photo_size,TeamID,isencomment) values ("&oblog.l_uid&",'"&ShowName&"','"&F_FileName&"',"&sysclass&","&subject&","&fileID&",'"&left(photo_Readme,240)&"',"&ishide&",'"&photo_title&"',"&File.FileSize&","&TeamID&","&isencomment&")")
						'�����û���������Ƭ��Ŀ
						'If subject > 0 Then oblog.Execute ("UPDATE oblog_subject SET subjectlognum = subjectlognum + 1,photo_path='"&F_FileName&"' WHERE subjecttype=1 AND subjectid="&subject&" AND userid="&oblog.l_uid)
						If subject > 0 Then oblog.Execute ("UPDATE oblog_subject SET subjectlognum = subjectlognum + 1 WHERE subjecttype=1 AND subjectid="&subject&" AND userid="&oblog.l_uid)
					End If
				End If
				If tMode = "9" Then
					ShowName = "�û�ͷ��"
				ElseIf tMode = "8" Then
					ShowName = "Ⱥ��LOGO"
				End if
				If re<>"no" Then
					select Case file.filetype
					'����ϴ��ļ�����ΪͼƬ
						Case 1
							If Not WriteData Then
							'ͷ��user_setting.asp,Ⱥ��user_team.asp
								Response.Write "<script>parent.document.oblogform.ico.value='" & F_FileName & "';parent.getImg();</script>"
							Else
								If oblog.CacheConfig(67) = "1" Then
									'F_FileName���¸�ֵ
									F_FileName = "attachment.asp?path="&F_FileName
								End if
								If tMode="10" Then
									'�༭��ģʽ�ϴ��ļ�
									Response.Write "<script>parent.upload('" & F_FileName &"');</script>"
'									Response.Write "<script>parent.oblogform.log_files.value+='," & FileID & "';</script>"
								Else
									If C_Editor_Type=1 Then	Response.Write "<script>parent.oEdit1.putHTML(parent.oEdit1.getHTMLBody()+'<img src="""&F_FileName&"""/><br/>');</script>"
									If C_Editor_Type=2 Then Response.Write "<script>parent.oblog_InsertSymbol('<img src=" &F_FileName& "><br />\n');</script>"
									Response.Write "<script>parent.oblogform.log_files.value+='," & FileID & "';</script>"
								End If
							End if
						Case Else
							'F_FileName���¸�ֵ
							'Response.Write "<script>parent.oblogform.log_files.value+='," & FileID & "';</script>"

							'F_FileName = "attachment.asp?FileID="&fileID
							If tMode = "10" Then
								Response.Write "<script>parent.upload('" & F_FileName &"');</script>"
							Else
								F_FileName = "attachment.asp?FileID="&fileID
								If C_Editor_Type=1 Then Response.Write "<script>parent.oEdit1.putHTML(parent.oEdit1.getHTMLBody()+'<a href="""&F_FileName&""">" & ShowName & "</a><br/>');</script>"
								If C_Editor_Type=2 Then Response.Write "<script>parent.oblog_InsertSymbol('<a href=" &F_FileName& ">"&ShowName&"</a>');</script>"
							End If
					End select
				End If
				If WriteData Then
					oblog.execute("update oblog_user set user_upfiles_num=user_upfiles_num+1,user_upfiles_size=user_upfiles_size+"&File.FileSize&" where userid="&oblog.l_uid)
				Else
					'���ɾ����ǰ�ϴ���ͷ�����Ⱥ��LOGO
					CheckFileExist (F_FileName)
					'�������ݿ�
					If tMode = "9" Then
						oblog.Execute ("UPDATE oblog_user SET user_Icon1 = '"&F_FileName&"' WHERE userid = "&oblog.l_uid)
					ElseIf tMode = "8" Then
						oblog.Execute ("UPDATE oblog_team SET t_ico = '"&F_FileName&"'  WHERE TeamID ="&TeamID)
					End if
				End If
				Session ("CheckUserLogined_"&oblog.l_uName) = ""
				Oblog.CheckUserLogined()
				If Isphoto = 1 Then
					oblog.ShowMsg "�ϴ�ͼƬ�ɹ�",""
				Else
					oblog.ShowMsg "�ϴ��ɹ�!","upload.asp?re="&re&"&isphoto="&isphoto&"&tMode=" & tMode &"&TeamID="&TeamID
				End if
				Set File = Nothing
			Next
		Else
			oblog.ShowMsg "��ѡ��Ҫ�ϴ����ļ�","upload.asp?re="&re&"&isphoto="&isphoto&"&tMode=" & tMode &"&TeamID="&TeamID
			Exit Sub
		End If
	Set Upload = Nothing
End Sub
'����ϴ�Ŀ¼������Ŀ¼���Զ�����
Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	if upload_dir<>"" then
		uploadpath = Year(Date) & "-" & Month(Date)
	else
		uploadpath=""
	end If
	If Not WriteData Then uploadpath = ""
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	On Error Resume Next
	Set objFSO = Server.CreateObject(oblog.CacheCompont(1))
	if upload_dir<>"" then '����ϴ�Ŀ¼
		If objFSO.FolderExists(Server.MapPath(upload_dir))=False Then
			objFSO.CreateFolder Server.MapPath(upload_dir)
		End If
	end if
	If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
		objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
	End If
	If Err.Number = 0 And upload_dir<>"" And WriteData Then
		CreatePath = PathValue & uploadpath & "/"
	Else
		CreatePath = PathValue
	End If
	Set objFSO = Nothing
End Function
'��ѯ��ǰ�ϴ�Ⱥ��ͼƬ���û��Ƿ�Ϊ����Ա
Function GroupManageID()
	GroupManageID = False
	If oblog.CheckAdmin(0) Then
		GroupManageID = True
		Exit Function
	End if
	Dim rsGroup
	Set rsGroup = oblog.Execute ("select TeamID FROM oblog_team WHERE TeamID="&TeamID&" AND managerid = "& oblog.l_uid)
	If Not rsGroup.EOF Then
		GroupManageID = True
	End If
	rsGroup.close
	Set rsGroup = Nothing
End Function
'����û��Ƿ��Ѿ��ϴ���ͷ�����Ⱥ��LOGO
Sub CheckFileExist(ByVal filepath)
	On Error Resume Next
	Dim objFSO,trs,tpath
	Set objFSO = Server.CreateObject(oblog.CacheCompont(1))
	If tMode = "8" Then
		Set trs = oblog.Execute ("select t_ico FROM oblog_team WHERE TeamID ="&TeamID )
		If Not trs.EOF Then
			tpath = trs(0)
		End If
		trs.close
		Set trs = Nothing
	Else
		tpath  = oblog.l_uIco
	End If
	'���ͼƬΪϵͳĬ��������
	If InStr (LCase(tpath),"images/") > 0 Then Exit Sub
	'����ϴ����ļ���ʽ��ͬ����Զ����ǣ�����Ҫɾ���������ɾ�����ϴ����ļ�
'	If tpath = blogdir&filepath Or tpath = filepath Then Exit Sub
	If tpath <> "" And Not IsNull(tpath) And Left(LCase(tpath),7)<>"http://" Then
		If objFSO.FileExists(Server.MapPath(tpath)) Then
			objFSO.DeleteFile Server.MapPath(tpath)
		End If
	End If
	Set objFSO = Nothing
End Sub
'����Ƿ�ΪȺ��ĳ�Ա
Function CheckQQMember()
	Dim rs
	CheckQQMember=False
	If oblog.checkuserlogined() Then
		Set rs=oblog.Execute("select id From oblog_teamusers Where state>2 and teamid=" & teamID & " And userid=" & oblog.l_uid )
		If Not rs.Eof Then
			CheckQQMember=True
		End If
		Set rs=Nothing
	Else
		If oblog.CheckAdmin(0) Then
			CheckQQMember = True
		End if
	End If
End Function
%>