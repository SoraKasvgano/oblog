<!--#include file="user_top.asp"-->
<script>
  function DivSetVisible(state)
  {
   var DivRef = document.getElementById('swin3');
   var IfrRef = document.getElementById('DivShim');
   if(state)
   {
    DivRef.style.display = "block";
    IfrRef.style.width = DivRef.offsetWidth;
    IfrRef.style.height = DivRef.offsetHeight;
    IfrRef.style.top = DivRef.style.top;
	//alert(DivRef.style.left);
    IfrRef.style.left = DivRef.style.left;
    IfrRef.style.zIndex = DivRef.style.zIndex - 1;
    IfrRef.style.display = "block";
   }
   else
   {
    DivRef.style.display = "none";
    IfrRef.style.display = "none";
   }
  }


</script>
<script src="inc/function.js" type="text/javascript"></script>
<%
Dim ssql,i,lPage,lAll,lPages,iPage,sGuide
iPage =30
Dim rs, sql, action
Dim id, cmd, Keyword, sField
Keyword = Trim(Request("keyword"))
If Keyword <> "" Then
    Keyword = oblog.filt_badstr(Keyword)
End If
sField = Trim(Request("Field"))
cmd = Trim(Request("cmd"))
action = Trim(Request("action"))
id = oblog.filt_badstr(Trim(Request("id")))
If cmd = "" Then
    cmd = 0
Else
    cmd = Int(cmd)
End If
G_P_FileName = "user_files.asp?cmd=" & cmd & "&page="
'此处组织纪录集
 ssql = "userid,file_name,file_path,file_size,fileid,file_readme,file_ext,isphoto,logid,file_showname,addtime"
    select Case cmd
        Case 0
            sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.l_uid&" order by fileid desc"
            sGuide = sGuide & "所有文件"
        Case 1
            sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.l_uid&" and ( file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' ) order by fileid desc"
'			sql="select "&ssql&" from oblog_upfile where userid="&oblog.l_uid&" and FileType=1 order by fileid desc"
            sGuide = sGuide & "图片文件"
        Case 2
			sql="select "&ssql&" from oblog_upfile where userid="&oblog.l_uid&" and FileType=2 order by fileid desc"
            sGuide = sGuide & "FLASH文件"
        Case 3
            sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.l_uid&" and ( file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm') order by fileid desc"
'			sql="select "&ssql&" from oblog_upfile where userid="&oblog.l_uid&" and FileType=3 order by fileid desc"
            sGuide = sGuide & "音频文件"
        Case 4
			sql="select "&ssql&" from oblog_upfile where userid="&oblog.l_uid&" and FileType=4 order by fileid desc"
            sGuide = sGuide & "视频文件"
        Case 5
            sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.l_uid&" and ( file_ext='rar' or file_ext='zip' or file_ext='arj' or file_ext='sit') order by fileid desc"
'			sql="select "&ssql&" from oblog_upfile where userid="&oblog.l_uid&" and FileType=5 order by fileid desc"
            sGuide = sGuide & "压缩文件"
        Case 6
            sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.l_uid&" and ( file_ext='doc' or file_ext='xsl' or file_ext='txt') order by fileid desc"
'			sql="select "&ssql&" from oblog_upfile where userid="&oblog.l_uid&" and FileType=6 order by fileid desc"
            sGuide = sGuide & "文档文件"
		Case 999
			sql="select "&ssql&" from oblog_upfile where userid="&oblog.l_uid&" and FileType=0 order by fileid desc"
            sGuide = sGuide & "其他文件"
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
<table id="TableBody" class="UserFilesBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th colspan="2">
				<ul id="UserMenu">
					<li><a href="#" onclick="chk_idAll(myform,1);">全部选择</a></li>
					<li><a href="#" onclick="chk_idAll(myform,0);">全部取消</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'删除选中的文件吗?')==true) {document.myform.submit();}">删除文件</a></li>
					<li><a href="#" onClick="return doMenu('swin4');">上传文件</a></li>
					<li id="showpage">
						<%If lAll>0 Then Response.Write MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th class="left"></th>
			<th>
				<table id="FilesTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"><%=sGuide%></td>
						<td class="t3">大小</td>
						<td class="t4">时间</td>
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
					<form name="myform1" method="post" action="user_files.asp">
						<select size=1 name='cmd' onChange='javascript:submit()'>
							<option value="10" selected="selected">请选择文件类型</option>
							<option value="0">列出所有文件</option>
							<option value="1">图片文件</option>
							<option value="2">FLASH</option>
							<option value="3">音频文件</option>
							<option value="4">视频文件</option>
							<option value="5">压缩照片</option>
							<option value="6">文档照片</option>
							<option value="999">其他照片</option>
						</select>
					</form>
					<br />
					<div id="space">
						<table cellpadding="0" title="使用空间：<%=oblog.showsize(oblog.l_uUpUsed)%>
剩余空间：<%=freesize%>
空间大小：<%=maxsize%>">
							<tr>
								<td class="used" width="<%=thispercent%>%" height="12"></td>
								<td width="100%"></td>
							</tr>
						</table>
						<ul>
							<li>使用空间：<span class="red"><%=oblog.showsize(oblog.l_uUpUsed)%></span></li>
							<li>剩余空间：<span class="red"><%=freesize%></span></li>
							<li>空间大小：<span class="red"><%=maxsize%></span></li>
						</ul>
					</div>
				</div>
			</td>
			<td>
				<div id="chk_idAll">
					<%
					Select Case action
						Case "modifyphoto"
							Call modify
						Case "savemodify"
							Call savemodify
						Case "delfile"
							Call delfile
						Case Else
							Call main()
					End Select
					Set rs = Nothing
					%>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
				<div id="swin1"></div>
				<div id="swin2"></div>
				<div id="swin3"></div>
				<div id="swin4" style="display:none;position:absolute;top:34px;left:259px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>上传文件</td>
						</tr>
						<tr>
							<td><iframe id='d_file' frameborder='0' src='upload.asp?tMode=<%=t%>&re=' width='100%' height='60' scrolling='no'></iframe></td>
						</tr>
						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin3');" value=" 确 定 " title=" 确 定 " /></td>
						</tr>
					</table>
				</div>
				<div id="swin5"></div>
				<iframe id="DivShim" scrolling="no" frameborder="0" style="position:absolute;top:0px; left:0px;display:none"></iframe>
</body>
</html>
<%
Sub main()
    Dim ext,imgsrc,imgsrc0
%>
					<form name="myform" method="post" action="user_files.asp?action=delfile" onSubmit="return confirm('确定要删除选定的文件吗？');">
					<table id="Files" class="TableList" cellpadding="0">
						<%
						i=0
						Do while not rs.eof
							imgsrc = rs("file_path")
							ext=rs("file_ext")
							If InStr("jpg,jpeg,gif,bmp,png,psd",ext) Then
								imgsrc0 = imgsrc
							Else
								imgsrc0 = "images/nopic.gIf"
							End if
						%>
						<tr id="u<%=cstr(rs("fileid"))%>" onclick="chk_iddiv('<%=cstr(rs("fileid"))%>')">
							<td class="t1" title="点击选中">
								<input name="id" type="checkbox" id="c<%=cstr(rs("fileid"))%>" value="<%=cstr(rs("fileid"))%>" onclick="chk_iddiv('<%=cstr(rs("fileid"))%>')">
							</td>
							<td class="t2"><!--<%=showfilepic(ext,rs("isphoto"))%>-->
								<a href="<%=imgsrc%>" onclick="chk_iddiv('<%=cstr(rs("fileid"))%>')" target="_blank" title="cssbody=[dogvdvbdy] cssheader=[dogvdvhdr] body=[<table cellpadding='0'><tr><td><img src='<%=imgsrc0%>' onload='javascript:if(this.width>190){this.resized=true;this.style.width=190;}' /></td></tr></table>] fixedabsx=[5] fixedabsy=[47]"><%=OB_IIF(rs("file_showname"),rs("file_name"))%></a>
							</td>
							<td class="t3">
								<%=oblog.showsize(rs("file_size"))%>
							</td>
							<td class="t4">
								<%=FormatDateTime(rs("addtime"),2)%>
							</td>
							<td class="t5">
								<a href="user_files.asp?action=delfile&id=<%=rs("fileid")%>" onclick="return confirm('确定要删除这个文件吗？');"><span class="red">删除</span></a>
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
<%
 rs.Close
    Set rs = Nothing
End Sub

function showfilepic(ext,isPhoto)
	Dim sReturn,sPhoto
	ext=lcase(ext)
	If isPhoto=1 Then
		sPhoto=",相册文件"
	Else
		sPhoto=""
	End If
	select Case ext
		Case "jpg","jpeg"
			sReturn="<img src=""images/filetype/jpg.gif"" class=""fileimg"" alt=""JPG文件"&sPhoto&""" />"
		Case "gif"
			sReturn="<img src=""images/filetype/gif.gif"" class=""fileimg"" alt=""GIF文件"&sPhoto&""" />"
		Case "bmp"
			sReturn="<img src=""images/filetype/bmp.gif"" class=""fileimg"" alt=""BMP文件"&sPhoto&""" />"
		Case "png"
			sReturn="<img src=""images/filetype/png.gif"" class=""fileimg"" alt=""PNG文件"&sPhoto&""" />"
		Case "psd"
			sReturn="<img src=""images/filetype/psd.gif"" class=""fileimg"" alt=""PSD文件"" />"
		Case "rar" ,"zip","arj","sit"
			sReturn="<img src=""images/filetype/rar.gif"" class=""fileimg"" alt=""压缩文件"" />"
		Case "xsl"
			sReturn="<img src=""images/filetype/excel.gif"" class=""fileimg"" alt=""Excel文件"" />"
		Case "doc"
			sReturn="<img src=""images/filetype/word.gif"" class=""fileimg"" alt=""Word文件"" />"
		Case "mp3"
			sReturn="<img src=""images/filetype/mp3.gif"" class=""fileimg"" alt=""mp3文件"" />"
		Case "rm","ram"
			sReturn="<img src=""images/filetype/rm.gif"" class=""fileimg"" alt=""Real文件"" />"
		Case "wmv" ,"wma","mpg" ,"avi"
			sReturn="<img src=""images/filetype/media.gif"" class=""fileimg"" alt=""媒体文件"" />"
		Case else
			sReturn="<img src=""images/filetype/blank.gif"" class=""fileimg"" alt=""其他文件"" />"
	end select
'	If InStr("jpg,jpeg,gif,bmp,png,psd",ext) Then sReturn="<img src="""&filepath&""" width=64 height=64 alt=""图片"&sPhoto&"""/>"
	showfilepic=sReturn
end function

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
    oblog.ShowMsg "删除文件成功！", ""
End Sub

Sub delonefile(id)
	id = CLng(id)
	Dim userid, filesize, filepath, fso, isphoto, imgsrc
	sql="select userid ,file_size,file_path,isphoto from [oblog_upfile] where fileid=" & id&" and userid="&oblog.l_uid
	Set rs = Server.CreateObject("adodb.recordset")
	rs.open sql, conn, 1, 3
	If Not rs.EOF Then
		userid = rs("userid")
		filesize = CLng(rs("file_size"))
		filepath = rs("file_path")
		isphoto = rs("isphoto")
		rs.Delete
		rs.Update
		rs.Close
		oblog.execute("update [oblog_user] set user_upfiles_num=user_upfiles_num-1,user_upfiles_size=user_upfiles_size-"&filesize&" where userid="&userid)
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
	If isphoto = 1 Then
		Set rs = oblog.Execute ("SELECT COUNT(commentid) FROM oblog_albumcomment WHERE mainid="&id)
		oblog.execute ("update [oblog_user] set comment_count = comment_count -"&OB_IIF(rs(0),0)&" where userid="&oblog.l_uid)
		rs.close
		oblog.Execute ("DELETE FROM oblog_album WHERE fileid = "&id)
		oblog.execute ("DELETE FROM [oblog_albumcomment] WHERE mainid = "&id)
	End  if
	Set rs = Nothing
End Sub
%>