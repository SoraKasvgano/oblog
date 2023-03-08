<!--#include file="user_top.asp"-->
<script src="inc/function.js" type="text/javascript"></script>
<script>
function chkmove(){
	var mid=read_checkbox('id');
	if (mid==''){
		alert("请选择要移动的相片");
		return false;
	}else{
		document.getElementById('moveid').value=mid;
		return true;
	}
}

</script>

<%
If oblog.CacheConfig(76) = "0" Then
	oblog.adderrstr("此功能已被系统关闭！")
	oblog.showusererr
End if
Dim ssql,i,lPage,lAll,lPages,iPage,sGuide
iPage =20
Dim rs, sql, action
Dim id, cmd, Keyword, sField,subjectid
Keyword = Trim(Request("keyword"))
If Keyword <> "" Then
    Keyword = oblog.filt_badstr(Keyword)
End If
sField = Trim(Request("Field"))
cmd = Trim(Request("cmd"))
action = Trim(Request("action"))
id = oblog.filt_badstr(Trim(Request("id")))
subjectid=Trim(Request("subjectid"))
If cmd = "" Then
	cmd = 0
Else
	cmd = Int(cmd)
End If
If subjectid = "" Then
	subjectid = 0
Else
	subjectid = clng(subjectid)

End If
G_P_FileName = "user_photo.asp?cmd=" & cmd & "&subjectid=" & subjectid & "&page="
'此处组织纪录集
ssql = "userid,photo_name,photo_path,photo_size,fileid,photo_readme,addtime,photo_title,commentnum,sysclassid"
	select Case cmd
		Case 0
			sql="select "&ssql&" from [oblog_album] where userid="&oblog.l_uid&" order by photoID desc"
			sGuide = sGuide & "所有相片"
		Case 1
			sql="select "&ssql&" from [oblog_album] where userid="&oblog.l_uid&" AND userClassId="&subjectid&" order by photoID desc"
			sGuide = sGuide & "分类图片"
		Case Else
	End select
	Set rs = Server.CreateObject("Adodb.RecordSet")
	rs.open sql, conn, 1, 3
	lAll=Int(rs.recordcount)
	'分页
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = Int(Request("page"))
	End If

	'设置缓存大小 = 每页需显示的记录数目
	rs.CacheSize = iPage
	rs.PageSize = iPage
	If lAll>0 Then
		rs.movefirst
		lPages = rs.PageCount
		If lPage>lPages Then lPage=lPages
		rs.AbsolutePage = lPage
	End If
'在后面进行实际的内容显示
%>
<%
select Case action
	Case "modifyphoto"
		Call modifyphoto
	Case "savemodify"
		Call savemodify
	Case "delfile"
		Call delfile
	case "movephoto"
		call movephoto()
	Case "isdefault"
		Call setdefault()
	Case Else
		Call main()
End select
Set rs = Nothing
%>
</body>
</html>
<%

Sub main()
%>
<table id="TableBody" class="UserFilesBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th colspan="2">
				<ul id="UserMenu">
					<li><a href="#" onclick="chk_idAll(myform,1);">全部选择</a></li>
					<li><a href="#" onclick="chk_idAll(myform,0);">全部取消</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'删除选中的文件吗?')==true) {document.myform.submit();}">删除文件</a></li>
					<li><a href="#" onClick="return doMenu('swin3');">移动分类</a></li>
					<li id="showpage">
						<%If lAll>0 Then Response.Write MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th class="left"></th>
			<th>
				<table id="PhotoTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"></td>
						<td class="t3"><%=sGuide%></td>
						<td class="t4">评</td>
						<td class="t5">操作</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td class="left">
				<%
				   Dim freesize, maxsize,maxsize1,thisPercent
						maxsize1 = oblog.l_Group(24,0)
						If maxsize1>0 Then
							maxsize = oblog.showsize(maxsize1 * 1024)
							freesize = oblog.showsize(Int(maxsize1*1024 - oblog.l_uUpUsed))
								thisPercent=oblog.l_uUpUsed/(maxsize1*1024)*100
						Elseif maxsize1=0 Then
							maxsize = "不限"
							freesize = "不限"
							thisPercent=0
						Elseif maxsize1=-1 Then
							maxsize = 0
							freesize = 0
							thisPercent=100
						End If
				%>
				<div id="viewimg"></div>
				<div id="content">
					<form name="myform1" method="post" action="user_photo.asp">
						<select size=1 name='subjectid' onChange='javascript:submit()'>
							<%
							dim substr,rst,tstr
							Response.Write "<option value=''>请选择相片分类</option>"
							Set rst = oblog.Execute("select subjectid,subjectname,ishide from oblog_subject where userid=" & oblog.l_uId & " And subjecttype=1")
							While Not rst.EOF
								If rst(2) = 1 Then tstr = "(隐藏)"
								substr=substr&"<option value="&rst(0)&">"&rst(1)&tstr&"</option>"
								tstr = ""
								rst.Movenext
							Wend
							Response.Write (substr)
							Response.Write "<option value=0>未分类</option>"
							set rst=nothing
							%>
						</select>
						<input type="hidden" value="1" name="cmd" />
					</form>
					<br />
					<div id="space">
						<table cellpadding="0" title="使用空间：<%=oblog.showsize(oblog.l_uUpUsed)%>
剩余空间：<%=freesize%>">
							<tr>
								<td class="used" width="<%=thispercent%>%" height="12"></td>
								<td width="100%"></td>
							</tr>
						</table>
						<ul>
							<li>使用空间：<span class="red"><%=oblog.showsize(oblog.l_uUpUsed)%></span></li>
							<li>剩余空间：<span class="red"><%=freesize%></span></li>
						</ul>
					</div>
				</div>
			</td>
			<td>
				<div id="chk_idAll">
					<form name="myform" method="post" action="user_photo.asp?action=delfile" onSubmit="return confirm('确定要删除选定的相片吗？');">
					<table id="Photo" class="TableList" cellpadding="0">
						<%
						i=0
						Do while not rs.eof
						%>
						<tr id="u<%=cstr(rs("fileid"))%>" onclick="chk_iddiv('<%=cstr(rs("fileid"))%>')">
							<td class="t1" title="点击选中">
								<input name="id" type="checkbox" id="c<%=cstr(rs("fileid"))%>" value="<%=cstr(rs("fileid"))%>" onclick="chk_iddiv('<%=cstr(rs("fileid"))%>')">
							</td>
							<td class="t2">
								<%
		Response.Write oblog.GetClassName(2,1,rs("sysclassid"))
		%>
							</td>
							<td class="t3">
								<a href="go.asp?fileid=<%=cstr(rs("fileid"))%>" onclick="chk_iddiv('<%=cstr(rs("fileid"))%>')" target="_blank" title="cssbody=[dogvdvbdy] cssheader=[dogvdvhdr] body=[<table cellpadding='0'><tr><td><img src='<%=rs("photo_path")%>' onload='javascript:if(this.width>190){this.resized=true;this.style.width=190;}' /></td></tr></table>] fixedabsx=[5] fixedabsy=[47]"><%=OB_IIF(rs("photo_title"),"未命名")%></a><span class="red"><%=oblog.showsize(rs("photo_size"))%></span>
								<!--时间-->
								<div class="time"><%=FormatDateTime(rs("addtime"),0)%></div>
							</td>
							<td class="t4">
								<%=rs("commentnum")%>
							</td>
							<td class="t5">
							<a href="user_photo.asp?action=isdefault&id=<%=rs("fileid")%>" onclick="return confirm('确定要将此相片设为封面吗？');"><span >设为封面</span></a>
								<a href="user_photo.asp?action=modifyphoto&id=<%=rs("fileid")%>" ><span class="green">修改</span></a>
								<a href="user_photo.asp?action=delfile&id=<%=rs("fileid")%>" onclick="return confirm('确定要删除这个相片吗？');"><span class="red">删除</span></a>
							</td>
						</tr>
						<%
							i = i + 1
							If i >= iPage Then Exit Do
							rs.movenext
						Loop
						%>
					</table>
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<div id="swin1" style="display:none;position:absolute;top:41px;left:10px;z-index:100;"></div>
<div id="swin2" style="display:none;position:absolute;top:41px;left:10px;z-index:100;"></div>
<div id="swin3" style="display:none;position:absolute;top:34px;left:259px;z-index:100;">
<form name="movesub" aciton="user_photo.asp" onSubmit="return chkmove();" method="post">
	<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
		<tr>
			<td align='center' class='win_table_top'>将选定的相片移动到分类</td>
		</tr>
		<tr>
			<td>
			目标分类：<select size="1" name='movesubjectid' >
			<%=substr%>
			</select>
			</td>
		</tr>
		<tr>
			<td class="win_table_end">
			<input type="hidden" name="moveid" value=""/>
			<input type="hidden" name="action" value="movephoto"/>
			<input type="submit" value=" 移动 ">　<input type="button" onClick="return doMenu('swin4');" value=" 关闭 " title="关闭" /> </td>
		</tr>
	</table>
</form>
</div>
<div id="swin4" style="display:none;position:absolute;top:41px;left:10px;z-index:100;"></div>
<div id="swin5" style="display:none;position:absolute;top:41px;left:10px;z-index:100;"></div>
<iframe id="DivShim" scrolling="no" frameborder="0" style="position:absolute;top:0px; left:0px;display:none"></iframe>
<%
 rs.Close
    Set rs = Nothing
End Sub

Sub delfile()
    If id = "" Then
        oblog.adderrstr ("错误：请指定要删除的文件！")
        oblog.showusererr
        Exit Sub
    End If

    If InStr(id, ",") > 0 Then
        Dim n, i
        id = FilterIDs(id)
	    n = Split(id, ",")
        For i = 0 To UBound(n)
            delonefile (n(i))
        Next
    Else
        delonefile (id)
    End If
    Set rs = Nothing
	oblog.reset_album_cover oblog.l_uid,"0"
    oblog.ShowMsg "删除相片成功！", ""
End Sub

Sub delonefile(id)
On Error Resume Next
    id = CLng(id)
    Dim userid, filesize, filepath, fso, isphoto, imgsrc,fileid
    sql="select userid ,file_size,file_path,isphoto,fileid from [oblog_upfile] where fileid=" & id&" and userid="&oblog.l_uid
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql, conn, 1, 3
    If Not rs.EOF Then
        userid = rs("userid")
        filesize = CLng(rs("file_size"))
        filepath = rs("file_path")
        isphoto = rs("isphoto")
		fileid = rs("fileid")
        rs.Delete
        rs.Update
        rs.Close
		Set rs = Nothing
		oblog.Execute ("delete from [oblog_Album] where fileid=" & fileid)
		oblog.Execute ("delete from [oblog_AlbumComment] where mainid=" & fileid)
		Set rs = oblog.Execute ("SELECT COUNT(commentid) FROM oblog_albumcomment WHERE mainid="&fileid)
		oblog.execute ("update [oblog_user] set comment_count = comment_count -"&OB_IIF(rs(0),0)&" where userid="&userid)
		rs.close
		Set rs = Nothing

		oblog.execute("update [oblog_user] set user_upfiles_num=user_upfiles_num-1,user_upfiles_size=user_upfiles_size-"&filesize&" where userid="&oblog.l_uid)
        If filepath <> "" Then
            imgsrc = filepath
            Set fso = Server.CreateObject(oblog.CacheCompont(1))
            If InStr("jpg,bmp,gif,png,pcx", Right(imgsrc, 3)) > 0 Then
                imgsrc = Replace(imgsrc, Right(imgsrc, 3), "jpg")
                imgsrc = Replace(imgsrc, Right(imgsrc, Len(imgsrc) - InStrRev(imgsrc, "/")), "pre" & Right(imgsrc, Len(imgsrc) - InStrRev(imgsrc, "/")))
                If fso.FileExists(Server.MapPath(imgsrc)) Then
                    fso.DeleteFile Server.MapPath(imgsrc)
                End If
            End If
            If fso.FileExists(Server.MapPath(filepath)) Then
                fso.DeleteFile Server.MapPath(filepath)
            End If
            Set fso = Nothing
		End If
    Else
        rs.Close
    End If

End Sub
Sub setdefault()
	If Not IsNumeric(id) Then ob_debug "参数错误",1
	oblog.execute("update oblog_album set is_album_default = 9 where fileid="&id)
	oblog.reset_album_cover oblog.l_uid,id
	oblog.execute("update oblog_album set is_album_default = 1 where fileid="&id)
	oblog.ShowMsg "成功设为此相册默认封面",""
End Sub

sub movephoto()
	dim id,subjectid
	id=Trim(Request("moveid"))
    If id = "" Then
        oblog.adderrstr ("请指定要移动的相片")
        oblog.showusererr
        Exit Sub
    End If
    subjectid = Trim(Request("movesubjectid"))
    If subjectid = "" Then
        oblog.adderrstr ("请指定要移动的目标分类")
        oblog.showusererr
        Exit Sub
    Else
        subjectid = CLng(subjectid)
    End If
    If InStr(id, ",") > 0 Then
        id = FilterIDs(id)
        sql="Update [oblog_album] set userclassid="&subjectid&" where fileid in (" & id & ") and userid="&oblog.l_uid
    Else
        sql="Update [oblog_album] set userclassid="&subjectid&" where fileid=" & CLng(id) &" and userid="&oblog.l_uid
    End If
    oblog.Execute sql
	Dim rst,rsu
	set rst=Server.CreateObject("adodb.recordset")
	rst.open "select subjectid,subjectlognum,subjecttype from oblog_subject where subjecttype = 1 AND  userid="&oblog.l_uid,conn,2,2
	while not rst.eof
		Set rsu = oblog.Execute ("SELECT COUNT(photoid) FROM oblog_album WHERE ishide = 0 AND  userclassid = "&rst(0))
		if not rsu.eof then rst("subjectlognum")=rsu(0) else rst("subjectlognum")=0
		rst.update
		rst.movenext
	wend
	rst.close
	Set rst = Nothing
    Set rs = Nothing
	oblog.reset_album_cover oblog.l_uid,"0"
    oblog.ShowMsg "移动到目标分类成功!", "user_photo.asp?cmd=1&subjectid="&subjectid

end sub
%>
<%
sub modifyphoto()
	dim id,rs,sql,trs
	dim restr
	id=Trim(Request("id"))
	if id="" then
		oblog.adderrstr( "错误：参数不足！")
		oblog.showusererr
		exit sub
	else
		id=CLng(id)
	end if
	sql="select * from [oblog_album] where fileid=" & id&" and userid="&oblog.l_uid
	set rs=oblog.execute(sql)
	if rs.bof then
		rs.close
		set rs=nothing
		oblog.adderrstr( "错误：找不到指定的文件！")
		oblog.showusererr
		exit sub
	end if
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="Photo" class="FieldsetForm">
						<legend>修改相片：<%If rs("teamid") > 0 Then Response.Write "("&oblog.CacheConfig(69)&"相片)"%></legend>
						<form action="user_photo.asp?action=savemodify" method="post" name="oblogform">
							<ul>
								<li><a href="<%=rs("photo_path")%>" title="点击查看原图" target="_blank"><%="<img src="""&rs("photo_path")&""" onload=""javascript:if(this.height>'190'){this.resized=true;this.style.height=190+'px';}"" />"%></a></li>
								<li>
									<label>相片标题：
										<input type="text" name = "photo_title" id="photo_title" size="50" maxlength="25" value="<%=rs("photo_title")%>" />
									</label>
								</li>
								<li>系统分类：
									<select name="photoclass">
										<%=oblog.show_class("log",rs("sysclassid"),1)%>
									</select>
								</li>
							<%If rs("TeamID") = 0 Then %>
								<li>相片分类：
									<select name="subjectid">
										<option value="0">我的分类</option>
										<%
										Set trs = oblog.Execute("select subjectid,subjectname from oblog_subject where userid=" & oblog.l_uid & " And subjectType=1")
										While Not trs.EOF
											If trs(0) = rs("userclassid") Then
												Response.Write ("<option value=" & trs(0) & " selected>" & oblog.filt_html(trs(1)) & "</option>")
											Else
												Response.Write ("<option value=" & trs(0) & " >" & oblog.filt_html(trs(1)) & "</option>")
											End If
											trs.movenext
										Wend
										Set trs = Nothing
										%>
									</select>
								</li>
								<li>是否隐藏：
									<label><input type="radio" name="ishide" id="ishide" value="0"  <%If rs("ishide") =0 Then %>Checked <%End if%>/>否</label>&nbsp;
									<label><input type="radio" id="ishide" name="ishide" value="1" <%If rs("ishide") =1 Then %>Checked <%End if%>/>是</label>
								</li>
							<%End if%>
								<li>评论开关：
									<label><input type="radio" id="isencomment" name="isencomment" value="1" <%If rs("isencomment") =1 Then %>Checked <%End if%>/>开</label>&nbsp;
									<label><input type="radio" name="isencomment" id="isencomment" value="0" <%If rs("isencomment") =0 Then %>Checked <%End if%>/>关</label>
								</li>
								<li>是否相册：
									<label><input type="radio" id="isphoto" name="isphoto" value="1" checked/>是</label>&nbsp;
									<label><input type="radio" name="isphoto" id="isphoto" value="0" />否</label>
									<font color="red">(如果选否,则只能在<a href="user_files.asp">文件管理</a>中找到此相片)</font>
								</li>
								<li>
									<label>
										相片说明：（500字内）<br />
										<textarea name="photo_readme" cols="45" rows="5"><%=oblog.filt_html(rs("photo_readme"))%></textarea>
									</label>
								</li>
								<li>
									<input type="hidden" name="id" value="<%=rs("fileid")%>" />
									<input type="submit" id="Submit" value="保存修改" />
								</li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
	rs.close
	set rs=nothing
end sub

sub savemodify()
	dim id,rs,sql
	Dim photo_title,photo_readme,isphoto
	id=Request("id")
	isphoto=CLng(Request("isphoto"))
	photo_title=Left(Trim(Request("photo_title")),30)
	photo_readme=Left(Trim(Request("photo_readme")),500)
	if id="" then
		oblog.adderrstr( "错误：参数不足！")
		oblog.showusererr
		exit sub
	Else
		id=CLng(id)
	end If
	If isphoto = 0 Then
		oblog.Execute ("DELETE FROM oblog_album WHERE fileid=" &id&" and userid="&oblog.l_uid)
		oblog.Execute ("UPDATE oblog_upfile SET isphoto = 0 WHERE fileid=" &id&" and userid="&oblog.l_uid)
		oblog.ShowMsg "编辑成功","user_photo.asp"
	End if
	If photo_title = "" Then
		oblog.adderrstr( "错误：至少需要填写图片标题！")
		oblog.showusererr
		exit Sub
	End If
	If oblog.chk_badword(photo_title) > 0 Then
		oblog.adderrstr( "错误：相片标题含有系统不允许的字符！")
		oblog.showusererr
		exit Sub
	End If
	If oblog.chk_badword(photo_title) > 0 Then
		oblog.adderrstr( "错误：相片介绍含有系统不允许的字符！")
		oblog.showusererr
		exit Sub
	End If
	sql="select * from [oblog_album] where fileid=" & id&" and userid="&oblog.l_uid
	set rs=Server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("photo_title")=photo_title
		rs("photo_readme")=photo_readme
		rs("sysclassid")=Request("photoclass")
		rs("userclassid")=Request("subjectid")
		rs("ishide")=Request("ishide")
		rs("isencomment")=CLng(Request("isencomment"))
		rs.update
		rs.close
		set rs=nothing
	end If
	oblog.ShowMsg "编辑成功",""
end Sub
%>