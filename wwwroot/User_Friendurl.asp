<!--#include file="user_top.asp"-->
<%
Dim rs, sql, blog
Dim id, action
action = Trim(Request("action"))
id = CLng(Request("id"))
select Case action
    Case "addurl"
    Call addurl
    Case "del"
    Call delurl
    Case "modify"
    Call modifyurl
    Case "savemodi"
    Call savemodify
    Case "order"
    Call order
    Case Else
    Call main
End select
Set rs = Nothing
%>
</body>
</html>
<%


Sub addurl()
    Call uporder
    Dim urlname,url,logourl, rs, ordernum,urltype
    urlname = Trim(Request.Form("urlname"))
	url = Trim(Request.Form("url"))
	logourl = Trim(Request.Form("logourl"))
	urltype = cint(Trim(Request.Form("urltype")))
    If urlname = "" Or oblog.strLength(urlname) > 50 Then oblog.adderrstr ("友情连接名不能为空且不能大于50字符)！")
    If oblog.chk_badword(urlname) > 0 Then oblog.adderrstr ("友情连接名中含有系统不允许的字符！")
	If url="http://" or url="" Or oblog.strLength(url) > 250 Then oblog.adderrstr ("友情连接地址不能为空且不能大于250字符)！")
    If oblog.chk_badword(url) > 0 Then oblog.adderrstr ("友情连接地址中含有系统不允许的字符！")
	if urltype=1 then
		If logourl="http://" or logourl="" Or oblog.strLength(logourl) > 250 Then oblog.adderrstr ("图片连接地址不能为空且不能大于250字符)！")
		If oblog.chk_badword(url) > 0 Then oblog.adderrstr ("图片连接地址中含有系统不允许的字符！")
	end if
	if oblog.errstr<>"" then
		oblog.showusererr
		exit sub
	end if
    Set rs = oblog.execute("select max(ordernum) from oblog_friendurl where userid=" & oblog.l_uid)
    If Not IsNull(rs(0)) Then
        ordernum = rs(0) + 1
    Else
        ordernum = 1
    End If
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select top 1 * from [oblog_friendurl] Where 1=0", conn, 1, 3
    rs.addnew
    rs("urlname") = urlname
    rs("userid") = oblog.l_uid
    rs("ordernum") = ordernum
    rs("urltype") = urltype
	rs("logo")=logourl
	rs("urlname")=urlname
	rs("url")=url
    rs.Update
    rs.Close
    Set rs = Nothing
	update_links()
    oblog.ShowMsg "添加友情连接成功!", "user_friendurl.asp?t=" & t
End Sub

Sub delurl()
    Dim id
    id = CLng(Request.QueryString("id"))
    oblog.execute("delete  from [oblog_friendurl] where urlid="&id&" and userid="&oblog.l_uid)
    Call uporder
	update_links()
    oblog.ShowMsg "删除友情连接成功!", ""
End Sub

Sub savemodify()
    Dim urlname,url,logourl, rs, ordernum,urltype
    urlname = Trim(Request.Form("urlname"))
	url = Trim(Request.Form("url"))
	logourl = Trim(Request.Form("logourl"))
	urltype = cint(Trim(Request.Form("urltype")))
    If urlname = "" Or oblog.strLength(urlname) > 50 Then oblog.adderrstr ("友情连接名不能为空且不能大于50字符)！")
    If oblog.chk_badword(urlname) > 0 Then oblog.adderrstr ("友情连接名中含有系统不允许的字符！")
	If url="http://" or url="" Or oblog.strLength(url) > 250 Then oblog.adderrstr ("友情连接地址不能为空且不能大于250字符)！")
    If oblog.chk_badword(url) > 0 Then oblog.adderrstr ("友情连接地址中含有系统不允许的字符！")
	if urltype=1 then
		If logourl="http://" or logourl="" Or oblog.strLength(logourl) > 250 Then oblog.adderrstr ("图片连接地址不能为空且不能大于250字符)！")
		If oblog.chk_badword(url) > 0 Then oblog.adderrstr ("图片连接地址中含有系统不允许的字符！")
	end if
	if oblog.errstr<>"" then
		oblog.ShowMsg Replace(oblog.errstr,"_","\n"),""
		exit sub
	end if
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select * from [oblog_friendurl] where urlid="&id&" and userid="&oblog.l_uid,conn,1,3
    If Not rs.EOF Then
        rs("urlname") = urlname
		rs("userid") = oblog.l_uid
		rs("urltype") = urltype
		rs("logo")=logourl
		rs("urlname")=urlname
		rs("url")=url
		rs.update
    End If
    rs.Close
    Set rs = Nothing
	update_links()
    %>
    <script language="javascript">
    	parent.location.href="user_friendurl.asp";
  	</script>
    <%
End Sub

Sub order()
    Dim ordernum, modi, rs
    ordernum = CLng(Request.QueryString("ordernum"))
    modi = Request.QueryString("modi")
    select Case modi
        Case "up"
            If ordernum - 1 > 0 Then
                oblog.execute("update [oblog_friendurl] set ordernum=9999999 where ordernum="&ordernum-1&" and userid="&oblog.l_uid )
                oblog.execute("update [oblog_friendurl] set ordernum=ordernum-1 where ordernum="&ordernum&" and userid="&oblog.l_uid)
                oblog.execute("update [oblog_friendurl] set ordernum="&ordernum&" where ordernum=9999999"&" and userid="&oblog.l_uid)
            End If
        Case "down"
            Set rs = oblog.execute("select max(ordernum) from oblog_friendurl where userid=" & oblog.l_uid)
            If ordernum < rs(0) Then
                oblog.execute("update [oblog_friendurl] set ordernum=9999999 where ordernum="&ordernum+1&" and userid="&oblog.l_uid )
                oblog.execute("update [oblog_friendurl] set ordernum=ordernum+1 where ordernum="&ordernum&" and userid="&oblog.l_uid )
                oblog.execute("update [oblog_friendurl] set ordernum="&ordernum&" where ordernum=9999999 and userid=" & oblog.l_uid )
            End If
            Set rs = Nothing
    End select
    'uporder()
	update_links()
    Response.Redirect "user_friendurl.asp"
End Sub

Sub uporder()
    Dim rs, i, n
    n = 0
    Set rs = oblog.execute("select count(urlid) from [oblog_friendurl] where userid=" & oblog.l_uid )
    ReDim ordernum(rs(0))
    Set rs = oblog.execute("select urlid from [oblog_friendurl] where userid=" & oblog.l_uid &" order by ordernum")
    While Not rs.EOF
        ordernum(n) = rs(0)
        n = n + 1
        rs.movenext
    Wend
    i = 1
    For n = 0 To UBound(ordernum)
        oblog.execute("update oblog_friendurl set ordernum="&i&" where urlid="&CLng(ordernum(n)))
        i = i + 1
        'Response.Write(i)
    Next
    Set rs = Nothing
End Sub

Sub main()
%>
<script language="javascript">
function doMenu1(MenuName,URL){
	document.getElementById("chgClass").src=URL;
	document.getElementById(MenuName).style.display = "block";
	}
</script>
<table id="TableBody" class="Setting" cellpadding="0">
	<thead>
		<tr>
			<th>
				<ul id="TabPage2">
					<li id="left_tab1" class="Selected" onClick="javascript:border_left('TabPage2','left_tab1');" title="博客设置">博客设置</li>
					<li id="left_tab2" onClick="javascript:border_left('TabPage2','left_tab2');" title="博客设置">用户设置</li>
					<li id="left_tab3" onClick="javascript:border_left('TabPage2','left_tab3');" title="博客设置">共同撰写</li>
				</ul>

				<div id="left_menu_cnt">
					<ul id="dleft_tab1" class="Selected" style="display:block;">
						<li id="now11"><a href="user_setting.asp?action=0&div=11" title="常规设置">常规设置</a></li>
						<li id="now12"><a href="user_setting.asp?action=placard&div=12" title="博客公告">博客公告</a></li>
						<li id="now13" class="Selected"><a href="user_friendurl.asp" title="博客友情连接">博客友情连接</a></li>
						<li id="now14"><a href="user_setting.asp?action=links&div=14" title="博客友情连接">高级编辑友情链接</a></li>
						<li id="now15"><a href="user_setting.asp?action=blogpassword&div=15" title="加密博客">加密博客</a></li>
						<li id="now16"><a href="user_setting.asp?action=blogstar&div=16" title="申请博客之星">申请博客之星</a></li>
					</ul>
					<ul id="dleft_tab2">
						<li id="now21"><a href="user_setting.asp?action=userinfo&div=21" title="个人资料">个人资料</a></li>
						<li id="now22"><a href="user_setting.asp?action=userpassword&div=22" title="密码设置">密码设置</a></li>
						<li id="now23"><a href="user_setting.asp?action=userpassword&div=23" title="密码保护">密码保护</a></li>
					</ul>
					<ul id="dleft_tab3">
						<li id="now31"><a href="user_setting.asp?action=blogteam&div=31" title="我管理的团队">我管理的团队</a></li>
						<li id="now32"><a href="user_setting.asp?action=blogteam&div=32" title="我管理的团队">我加入的团队</a></li>
						<li id="now33"><a href="user_setting.asp?action=blogteam&div=33" title="我管理的团队">邀请朋友加入</a></li>
					</ul>
				</div>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<div class="btu">
						<span><a href="#" onClick="return doMenu('swin1');">添加友情链接</a></span>
						<div>如果您想要更个性的友情链接，您还可以使用<a href="user_setting.asp?action=links&div=14">高级编辑</a>来添加友情连接。</div>
					</div>
					<table id="Friendurl" class="TableList" cellpadding="0">
<%
Dim rs
Set rs = oblog.execute("select * from oblog_friendurl where userid=" & oblog.l_uid & " order by ordernum")
While Not rs.EOF
%>
						<tr>
							<td class="t1">
								<%=rs("ordernum")%>
							</td>
							<td class="t2">
								<%if rs("urltype")=0 then Response.Write("") Else Response.Write("<img src="""&oblog.filt_html(rs("logo"))&""" />")%>
							</td>
							<td class="t3">
								<%="<a href='"&oblog.filt_html(rs("url"))&"' target=""_blank"" title="""&oblog.filt_html(rs("urlname"))&""">"&oblog.filt_html(rs("urlname"))&"</a>"%>
								<div class="url"><%=oblog.filt_html(rs("url"))%></div>
							</td>
							<td class="t4">
								<a onClick="return doMenu1('swin2','user_friendurl.asp?action=modify&id=<%=rs("urlid")%>&oldname=<%=oblog.htm2js(rs("urlname"),False)%>&t=<%=rs("urltype")%>');" href="#"><span class="green">修改</span></a>
								<a href="user_friendurl.asp?action=del&id=<%=rs("urlid")%>&t=<%=t%>" <%="onClick='return confirm(""确定要删除此友情连接吗(不可恢复)？"");'"%>><span class="red">删除</span></a>
								<a href="user_friendurl.asp?action=order&modi=up&ordernum=<%=rs("ordernum")%>&t=<%=t%>"><span class="blue">上移</span></a>
								<a href="user_friendurl.asp?action=order&modi=down&ordernum=<%=rs("ordernum")%>&t=<%=t%>"><span class="blue">下移</span></a>

							</td>
						</tr>
<%
rs.movenext
Wend
Set rs = Nothing
%>
					</table>
					</form>
				</div>
			</td>
		</tr>
	</tbody>
</table>
<div id="swin1" style="display:none;position:absolute;top:90px;left:10px;z-index:99999;">
	<form name="form1" method="post" action="user_friendurl.asp?action=addurl&t=<%=t%>">
	<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
		<tr>
			<td colspan='2' align='center' class='win_table_top'>添加友情连接</td>
		</tr>
		<tr>
			<td class='win_table_td'>类型：</td>
			<td><label><input type="radio" name="urlType" value="0"  onClick="document.getElementById('logourl').disabled='disabled';" checked>&nbsp;文字链接</label>&nbsp;&nbsp;<label><input name="urlType" type="radio" value="1"  onClick="document.getElementById('logourl').disabled='';" >&nbsp;图片链接</label></td>
		</tr>

		<tr>
			<td class='win_table_td'>友情链接名：</td>
			<td><input name="urlname" type="text" id="urlname" maxlength="50" value="" /></td>
		</tr>
		<tr>
			<td class='win_table_td'>链接地址：</td>
			<td><input name="url" type="text" id="url" maxlength="255" size="40" value="http://" /></td>
		</tr>
		<tr>
			<td class='win_table_td'>图片地址：</td>
			<td><input name="logourl" type="text" id="logourl" maxlength="255" size="40" disabled="disabled" value="http://" /></td>
		</tr>

		<tr>
			<td colspan='2' class="win_table_end"> <input type="submit" value=" 添 加 " /> <input type="button" onClick="return doMenu('swin1');" value="关闭" title=" 关 闭 " /></td>
		</tr>
	</table>
	</form>
</div>
<div id="swin2" style="display:none;position:absolute;top:42px;left:10px;z-index: 99999;">
　<iframe class="FrmID" id="chgClass"  style="width:440px;height:195px;" src="" frameborder="0" scrolling="auto" onunload="parent.location.href='user_friendurl.asp?t=<%=t%>'"></iframe>
</div>
<div id="swin3"></div>
<div id="swin4"></div>
<div id="swin5"></div>
<iframe id="DivShim" scrolling="no" frameborder="0" style="position:absolute;top:0px; left:0px;display:none"></iframe>
</div>
<%
End Sub

Sub modifyurl()
    Dim rs
    id = CLng(id)
    set rs=oblog.execute("select * from oblog_friendurl where urlid="&id&" and userid="&oblog.l_uid)
    If Not rs.EOF Then

%>
	<form name="form1" method="post" action="user_friendurl.asp?action=savemodi&id=<%=id%>">
	<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
		<tr>
			<td colspan='2' align='center' class='win_table_top'>修改友情连接</td>
		</tr>
		<tr>
			<td class='win_table_td'>类型：</td>
			<td><label><input type="radio" name="urlType" value="0"  onClick="document.getElementById('logourl').disabled='disabled';" <%if rs("urltype")=0 then Response.Write("checked")%>>&nbsp;文字链接</label>&nbsp;&nbsp;<label><input name="urlType" type="radio" value="1"  onClick="document.getElementById('logourl').disabled='';" <%if rs("urltype")=1 then Response.Write("checked")%>>&nbsp;图片链接</label></td>
		</tr>
		<tr>
			<td class='win_table_td'>友情链接名：</td>
			<td><input name="urlname" type="text" id="urlname" maxlength="50" value="<%=oblog.filt_html(rs("urlname"))%>" /></td>
		</tr>
		<tr>
			<td class='win_table_td'>链接地址：</td>
			<td><input name="url" type="text" id="url" maxlength="255" size="40" value="<%=oblog.filt_html(rs("url"))%>" /></td>
		</tr>
		<tr>
			<td class='win_table_td'>图片地址：</td>
			<td><input name="logourl" type="text" id="logourl" maxlength="255" size="40"  <%if rs("urltype")=0 then Response.Write("disabled='disabled'")%> value="<%if oblog.filt_html(rs("logo"))="" then Response.Write("http://") else Response.Write(rs("logo"))%>" /></td>
		</tr>
		<tr>
			<td colspan='2' class="win_table_end"> <input type="submit" value=" 修 改 " /> <input type="button" onClick="return parent.doMenu('swin2');" value=" 关 闭 " title="关闭" /></td>
		</tr>
	</table>
	</form>
<%
    Set rs = Nothing
    End If
End Sub

sub update_links()
	dim blog
	set blog=new class_blog
	blog.userid=oblog.l_uid
	blog.update_links oblog.l_uid
	set blog=nothing
end sub

%>