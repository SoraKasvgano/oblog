<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/md5.asp"-->
<%
'ͼƬ����
Const g_photo_title = 30
'ͼ����
Const g_photo_about = 500
dim oblog,action,userid
set oblog=new class_sys
oblog.autoupdate=false
oblog.start
Response.contentType="application/xml"
Response.Expires=0
Response.Write("<?xml version=""1.0"" encoding=""UTF-8""?>")
action=Trim(Request("action"))
userid=CLng(Request("userid"))
select case action
	case "menu"
		call menu()
	case "photo"
		call photo()
	case "sysclass"
		call sysclass()
	case "subject"
		call subject()
	case "chklogin"
		call chklogin()
	case "islogin"
		call islogin()
	case "delfile"
		call delfile()
	case "getname"
		call getUName()
	case "writeabout"
		call writeabout()
	case "addcomment"
		call addcomment()
end select

sub menu()
	dim rs,menustr,i
	set rs=oblog.execute("select subjectname,subjectid from oblog_subject where userid=" & userid & " And SubjectType=1  and (ishide=0 or ishide is null)   order by ordernum")
	while not rs.eof
		menustr=menustr&"<menu label="""&rs("subjectname")&""" subjectid="""&rs("subjectid")&""" />"
		rs.movenext
	wend
%>
<menu>
	<menu label="����ͼƬ" subjectid="0" />
	<%=menustr%>
	<menu label="��ͷ��" subjectid="-1" />
</menu>
<%
	set rs=nothing
end sub
'��ȡ����б�
sub photo()
	dim rs,ssql,subjectid,pstr
	subjectid=CLng(Request("subjectid"))
	if subjectid>0 then ssql=" and userClassId="&subjectid
	if subjectid=-1 then ssql=" and isBigHead=1"
	ssql = ssql & " AND ( ishide = 0 OR ishide IS NULL) "
	set rs=oblog.execute("select photo_title,photo_readme,photo_path,fileid from oblog_album where userid="&userid&ssql&" AND TeamID = 0 order by photoID desc")
	while not rs.eof
		pstr=pstr&"<p _name="""&rs(0)&""" _url="""&rs(2)&""" _fileid="""&rs(3)&""" _about="""&rs(1)&""" />"
		rs.movenext
	wend
	Response.Write "<photo>"&pstr&"</photo>"
	set rs=nothing
end sub
'��ȡϵͳ������
sub sysclass()
	dim rs,pstr
	set rs=oblog.execute("select id,classname from oblog_logclass where idType=1 order by RootID,OrderID")
	while not rs.eof
		pstr=pstr&"<c _name="""&rs("classname")&""" _id="""&rs("id")&"""/>"
		rs.movenext
	wend
	Response.Write "<sysclass>"&pstr&"</sysclass>"
	set rs=nothing
end sub
'��ȡ�û��Զ������ר��
sub subject()
	dim rs,pstr
	set rs=oblog.execute("select subjectid,subjectname from oblog_subject where subjectType=1 and userid="&userid&"  and (ishide=0 or ishide is null) order by Ordernum")
	while not rs.eof
		pstr=pstr&"<s _name="""&rs("subjectname")&""" _id="""&rs("subjectid")&"""/>"
		rs.movenext
	wend
	Response.Write "<subject>"&pstr&"</subject>"
	set rs=nothing
end sub
'��ȡ�û���ǰ��¼���û���
sub getUName()
	dim xmlstr,u_name,test,UserUrl
	if not oblog.checkuserlogined() then
		xmlstr="<m msg=""δ��¼"" login=""0"" />"
	Else
		If oblog.CacheConfig(5) = "1" Then
			If Left(oblog.l_udomain,8)="http://." Or Trim(oblog.l_udomain)="." Then
				UserUrl=oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext
			Else
				UserUrl="http://"&oblog.l_udomain
			End If
		Else
			UserUrl=oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext
		End If
		u_name = oblog.l_uName
		xmlstr="<m msg="""&u_name&""" login=""1"" />"
	end if
	Response.Write "<getname>"&xmlstr&"</getname>"
end sub
'��֤�û�Ȩ��
sub islogin()
	dim xmlstr
	if not oblog.checkuserlogined() then
		xmlstr="<m msg=""δ��¼"" login=""0"" />"
	else
		if oblog.l_ulevel=6 then
			xmlstr="<m msg=""�����˺�δͨ�����"" login=""0"" />"
		elseif oblog.l_uid<>userid then
			xmlstr="<m msg=""�����ǵ�ǰ�û�����Ȩ�޲���"" login=""0"" />"
		else
			'xmlstr="<m msg=""�ѵ�¼"" login=""1"" />"
			xmlstr="<m msg=""�ѵ�¼"" login=""1"" maxsize="""&Int(oblog.l_Group(24,0))&""" onesize="""&oblog.l_Group(23,0)&""" title_num="""&g_photo_title&""" about_num="""&g_photo_about&""" />"
		end if
	end if
	Response.Write "<islogin>"&xmlstr&"</islogin>"
end sub
'�ж��û��Ƿ��¼
sub chklogin()
	dim username,password,xmlstr
	username=oblog.filt_badstr(Trim(Request("username")))
	password=md5(Request("password"))
	oblog.ob_chklogin username,password,0
	if oblog.errstr<>"" then
		xmlstr="<m msg="""&oblog.errstr&""" login=""0"" />"
	else
		xmlstr="<m msg=""��¼�ɹ�"" login=""1"" />"
	end if
	Response.Write "<chklogin>"&xmlstr&"</chklogin>"
end sub
'��������⡢���
sub writeabout()
	dim pname,pabout,fileid,xmlstr,myuserid
	pname = ProtectSQL(Trim(Request("pname")))
	pabout = ProtectSQL(Trim(Request("pabout")))
	fileid=CLng(Trim(Request("fileid")))
	myuserid=CLng(Trim(Request("userid")))
	'��ʼ�������ַ���
	xmlstr="<m msg="""&oblog.checkuserlogined()&""" isadd=""0"" />"
	'δ��¼
	if not oblog.checkuserlogined() then
		xmlstr="<m msg=""δ��¼"" isadd=""0"" />"
	else
		if oblog.l_ulevel=6 then
			xmlstr="<m msg=""�����˺�δͨ�����"" isadd=""0"" />"
		elseif oblog.l_uid<>myuserid then
			xmlstr="<m msg=""���ǵ�ǰ�û�"" isadd=""0"" />"
		else
			'����ִ�����
			if len(pname)>g_photo_title then
				xmlstr="<m msg='ͼƬ�����������ܳ���:"&g_photo_title&"' isadd='2' />"
				Response.Write "<file>"&xmlstr&"</file>"
				exit sub
			elseif len(pabout)>g_photo_about then
				xmlstr="<m msg='ͼƬ����������ܳ���:"&g_photo_about&"' isadd='2' />"
				Response.Write "<file>"&xmlstr&"</file>"
				exit sub
			end if

			dim userid,sql,rs
			sql="select * from [oblog_album] where fileID="&fileid&" and userid="&oblog.l_uid
			set rs=Server.CreateObject("adodb.recordset")
			link_database
			rs.open sql,conn,1,3
			if not rs.eof then
				rs("photo_title") = RemoveHtml(keyWordReplace(pname))
				rs("photo_readme") = RemoveHtml(keyWordReplace(pabout))
				rs.update
				xmlstr="<m msg=""�����ɹ�"" isadd=""1"" />"
			End If
			rs.close
			set rs=nothing
		end if
	end if
	Response.Write "<file>"&xmlstr&"</file>"

end sub
'ɾ������ļ�
sub delfile()
	dim fileid,xmlstr
	fileid=Trim(Request("fileid"))
	if not oblog.checkuserlogined() then
		Response.Write("<delfile><m msg=""���¼��ִ��ɾ������"" isdel=""0"" /></delfile>")
		exit Sub
	end if
	 If InStr(fileid, ",") > 0 Then
        Dim n, i
        fileid = FilterIDs(fileid)
        n = Split(fileid, ",")
        For i = 0 To UBound(n)
            if delonefile (n(i)) then
				xmlstr="<m msg=""ɾ���ɹ�"" isdel=""1"" />"
			else
				xmlstr="<m msg=""ɾ��ʧ��"" isdel=""0"" />"
			end if
        Next
    Else
        if delonefile (fileid) then
			xmlstr="<m msg=""ɾ���ɹ�"" isdel=""1"" />"
		else
			xmlstr="<m msg=""ɾ��ʧ��"" isdel=""0"" />"
		end if
    End If
	Response.Write "<delfile>"&xmlstr&"</delfile>"
end sub
'ɾ����������ļ�
function delonefile(fileid)

	fileid=CLng(fileid)
	dim userid,filesize,filepath,fso,isphoto,imgsrc,sql,rs,fid
	sql="select * from [oblog_upfile] where fileid=" & fileid&" and userid="&oblog.l_uid
	set rs=Server.CreateObject("adodb.recordset")
	link_database
	rs.open sql,conn,1,3
	if not rs.eof then
		userid=rs("userid")
		filesize=CLng(rs("file_size"))
		filepath=rs("file_path")
		isphoto=rs("isphoto")
		fid=rs("fileid")
		rs.delete
		rs.update
		rs.close
		oblog.Execute ("delete from [oblog_Album] where fileid=" & fid)
		oblog.Execute ("delete from [oblog_AlbumComment] where fileid=" & fid)
		oblog.execute("update [oblog_user] set user_upfiles_size=user_upfiles_size-"&filesize&" where userid="&userid)
		if filepath<>"" then
			imgsrc=filepath
			Set fso = Server.CreateObject(oblog.CacheCompont(1))
			if instr("jpg,bmp,gif,png,pcx",right(imgsrc,3))>0 then
				imgsrc=Replace(imgsrc,right(imgsrc,3),"jpg")
				imgsrc=Replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
				if  fso.FileExists(Server.MapPath(imgsrc)) then
					fso.DeleteFile Server.MapPath(imgsrc)
				end if
			end if
			if fso.FileExists(Server.MapPath(filepath)) then
				fso.DeleteFile Server.MapPath(filepath)
			end if
			set fso=nothing
		end if
		delonefile=true
	else
		delonefile=false
		rs.close
		set rs=nothing
	End If
	oblog.Execute ("DELETE FROM oblog_album WHERE fileid = "&fileid)
end Function
'FLASH����������
sub addcomment()

	dim c_name, c_homepage, c_title, c_content, photo_id, xmlstr, success_str
	'//
	photo_id=CLng(Trim(Request("fileid")))
	c_name = Trim(Request("comName"))
	c_homepage = Trim(Request("comHomepage"))
	c_title=(Trim(Request("comTitle")))
	c_content=(Request("comContent"))
	'//
	success_str = "<m msg=""�����ɹ�!"" isadd=""1"" />"
	xmlstr = success_str
	'//
	'���ύ�����ݽ��кϷ��Լ��
	'//
	'//����
	if oblog.strLength(c_name)>20 then xmlstr="<m msg=""���ֲ��ܴ���20���ַ�!"" isadd=""2"" />"
	'//����
	if oblog.strLength(c_content)>Int(oblog.CacheConfig(35)) then xmlstr="<m msg=""�ظ����ݲ��ܴ���"&oblog.CacheConfig(35)&"Ӣ���ַ�!"" isadd=""2"" />"
	if oblog.chk_badword(c_content)>0 then xmlstr="<m msg=""�ظ������к���ϵͳ��������ַ�!"" isadd=""2"" />"
	'//����
	if oblog.strLength(c_title)>200 then xmlstr="<m msg=""�ظ����ⲻ�ܴ���200Ӣ���ַ�!"" isadd=""2"" />"
	if oblog.chk_badword(c_title)>0 then xmlstr="<m msg=""�ظ������к���ϵͳ��������ַ�!"" isadd=""2"" />"
	'//��ҳ
	if oblog.chk_badword(c_homepage)>0 then xmlstr="<m msg=""��ҳ��ַ�к���ϵͳ��������ַ�!"" isadd=""2"" />"

	if not(xmlstr=success_str) then
		Response.Write "<file>"&xmlstr&"</file>"
		exit sub
	end if

	'//
	'//���ݿ�д�����
	dim rs
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from oblog_album_comments WHERE 1 = 0 ",conn,2,2
	rs.addnew
	rs("mainid")=photo_id
	rs("comment_user")=EncodeJP(keyWordReplace(c_name))
	rs("homepage")=keyWordReplace(c_homepage)
	rs("commenttopic")=EncodeJP(keyWordReplace(oblog.InterceptStr(oblog.filt_badword(c_title),250)))
	rs("comment")=keyWordReplace(c_content)
	rs("addtime")=oblog.ServerDate(now())
	rs("addip")=oblog.userip
	If oblog.checkuserlogined() Then
		rs("isguest") = 0
	Else
		rs("isguest") = 1
	End if
	rs.update
	rs.close
	set rs=Nothing
	'//

	Response.Write "<file>"&xmlstr&"</file>"

end Sub
'���˲��������ַ�
function keyWordReplace(str)
	On Error Resume Next
	dim tmpStr
	tmpStr = str
	tmpStr = Replace(str, "35;", "&#35;")  '#'
	tmpStr = Replace(tmpStr, "38;", "&#38;")  '&'
	tmpStr = Replace(tmpStr, "58;", "&#58;")
	tmpStr = Replace(tmpStr, "60;", "&#60;")
	keyWordReplace = Replace(tmpStr, "62;", "&#62;")
end function
%>