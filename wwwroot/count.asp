<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%
dim action,id,rs,n,oblog
dim refreshLimitTime,timeStamp,fv
Dim ShowDigg
Dim cookies_name_count
cookies_name_count  = cookies_name & "_count"
action=Request.QueryString("action")
set oblog=new class_sys
oblog.autoupdate=False
oblog.start
Response.Buffer = True
'关闭统计功能，不允许关闭
If oblog.cacheConfig(12) = "0" And 1=2 Then
	select Case action
		Case "site"
			Response.Write "site_count.innerHTML=""-"";"
			Call DiggCshow()
		Case "log"
			Response.Write "document.write('-');"
		Case "code"
			Call comment_code
		Case "code31"
			Call comment_code31
		Case "logtb"
			Response.Write "ob_logreaded.innerHTML=""-"";"
			Response.Write "ob_tbnum.innerHTML=""-"";"
		Case "logtb31"
		Case "logs"
			'暂不处理
	End select
	Response.End
End If

ShowDigg = vbcrlf & "<div class=""digg_list"" style=""float: right; display:inline; margin: 0 10px 5px 0; width: 45px; height: 55px; background: url("&blogurl&"Images/digg.gif) no-repeat left top; text-align: center; "">" & vbcrlf
ShowDigg = ShowDigg & "	<div class=""digg_number"" style=""width:45px;padding: 10px 0 11px 0;font-size:18px;font-weight:600;color:#333;font-family:tahoma,Arial,Helvetica,sans-serif;line-height:1.0;"">$diggnum$</div>" & vbcrlf
ShowDigg = ShowDigg & "	<div class=""digg_submit"" style="" padding: 3px 0 0 6px;line-height:1.0;letter-spacing: 6px; ""><a href=""javascript:void(null)"" onclick=""diggit($logid$);"" style=""font-size:12px;line-height:1.0;"">$showmsg$</a></div>" & vbcrlf
ShowDigg = ShowDigg & "</div>" & vbcrlf

select Case action
	Case "site"
		Call site_count
		Call DiggCshow()
	Case "log"
		Call log_count
	Case "code" '兼容3.0
		Call comment_code
	Case "code31"
		Call comment_code31
	Case "logtb"	'兼容3.0版本的统计
		Call logtb_count("3.0")
	Case "logtb31"	'3.1版本的日志统计，增加(）输出
		Call logtb_count("3.1")
	Case "logs"
		Call logs_count
	Case "ping"
		Call ping(Request("logid"))
end select

sub site_count()
	id=CLng(Request.QueryString("id"))
	refreshLimitTime  =  Int(oblog.CacheConfig(31))
	if refreshLimitTime="" or isnull(refreshLimitTime) then
		refreshLimitTime=0
	end if
	if Request.cookies(cookies_name_count)("lastvisit_fresh_site"&id)="" then
		if cookies_domain<>"" then Response.Cookies(cookies_name_count).Domain=cookies_domain
		Response.Cookies(cookies_name_count).Path   =   blogdir
		Response.cookies(cookies_name_count)("lastvisit_fresh_site"&id)=Time()
		fv=true
	end if
	timeStamp=Time()
	if not IsObject(conn) then link_database
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select user_siterefu_num from oblog_user where userid="&id,conn,1,3
	n=rs(0)+1
	if (datediff("s",Request.cookies(cookies_name_count)("lastvisit_fresh_site"&id),timeStamp)>refreshLimitTime) or fv=true then
		rs(0)=n
		rs.update
		if cookies_domain<>"" then Response.Cookies(cookies_name_count).Domain=cookies_domain
		Response.Cookies(cookies_name_count).Path   =   blogdir
		Response.cookies(cookies_name_count)("lastvisit_fresh_site"&id)=timeStamp
	end if
	rs.close
	set rs=nothing
	'Response.Write "document.write('"&n&"');"
	Response.Write oblog.htm2js_div(n,"site_count")
end sub

sub log_count()
	id=CLng(Request.QueryString("id"))
	refreshLimitTime  =  Int(oblog.CacheConfig(31))
	if refreshLimitTime="" or isnull(refreshLimitTime) then
		refreshLimitTime=0
	end if
	if Request.cookies(cookies_name_count)("lastvisit_fresh_log"&id)="" then
		if cookies_domain<>"" then Response.Cookies(cookies_name_count).Domain=cookies_domain
		Response.Cookies(cookies_name_count).Path   =   blogdir
		Response.cookies(cookies_name_count)("lastvisit_fresh_log"&id)=time()
		fv=true
	end if
	timeStamp=time()
	if not IsObject(conn) then link_database
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select iis from oblog_log where logid="&id,conn,1,3
	n=rs(0)+1
	if (datediff("s",Request.cookies(cookies_name_count)("lastvisit_fresh_log"&id),timeStamp)>refreshLimitTime)  or fv=true then
		rs(0)=n
		rs.update
		if cookies_domain<>"" then Response.Cookies(cookies_name_count).Domain=cookies_domain
		Response.Cookies(cookies_name_count).Path   =   blogdir
		Response.cookies(cookies_name_count)("lastvisit_fresh_log"&id)=timeStamp
	end if
	rs.close
	set rs=nothing
	Response.Write "document.write('"&n&"');"
end sub

sub logtb_count(ver)
	id=CLng(Request.QueryString("id"))
	Dim tbn,diggn,digg,Qs
	refreshLimitTime  =  Int(oblog.CacheConfig(31))
	if refreshLimitTime="" or isnull(refreshLimitTime) then
		refreshLimitTime=0
	end if
	if Request.cookies(cookies_name_count)("lastvisit_fresh_log"&id)="" then
		if cookies_domain<>"" then Response.Cookies(cookies_name_count).Domain=cookies_domain
		Response.Cookies(cookies_name_count).Path   =   blogdir
		Response.cookies(cookies_name_count)("lastvisit_fresh_log"&id)=time()
		fv=true
	end if
	timeStamp=Time()
	if not IsObject(conn) then link_database
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select iis,trackbacknum,diggnum,isDigg,IsSpecial from oblog_log a INNER JOIN oblog_user b ON a.authorid = b.userid where logid="&id,conn,1,3
	n=rs(0)+1
	tbn=rs(1)
	If OB_IIF(rs("IsSpecial"),0) = 0 Then
	diggn = rs(2)
	digg = ShowDigg
	digg = Replace(digg,"$diggnum$",OB_IIF(diggn,0))
	digg = Replace(digg,"$logid$",id)
	digg = Replace(digg,"$showmsg$","推荐")
	End If
'	Qs = "|<ul class=""Reflect"" style=""display:inline;"" id=""menu_"&rs(0)&""" onmouseover=""menuFix('"&rs(0)&"');""> <li> <a href=""javascript:void(null)"">反映问题</a> <ul>"&GetMenuList(id)&"</ul>"
	If IsArray(oblog.CacheReport)Then
		If UBound(oblog.CacheReport) > 0 Then
			Qs = "|<a href=""javascript:void(null)"" onclick=""openScript('"&blogurl&"report.asp?logid="&id&"',450,400)"">反映问题</a> "
		End If
	End if
	if (DateDiff("s",Request.cookies(cookies_name_count)("lastvisit_fresh_log"&id),timeStamp)>refreshLimitTime)  or fv=true Then
		rs(0)=n
		rs.update
		if cookies_domain<>"" then Response.Cookies(cookies_name_count).Domain=cookies_domain
		Response.Cookies(cookies_name_count).Path   =   blogdir
		Response.cookies(cookies_name_count)("lastvisit_fresh_log"&id)=timeStamp
	end if
	if ver="3.0" then
		Response.Write "document.getElementById('ob_logreaded').innerHTML="""&n&""";"
		Response.Write "document.getElementById('ob_tbnum').innerHTML="""&tbn&""";"
	else
		Response.Write oblog.htm2js_div("("&n&")","ob_logreaded")
		Response.Write oblog.htm2js_div("("&tbn&")","ob_tbnum")
		If OB_IIF(rs("isdigg"),1) = 1 And OB_IIF(rs("IsSpecial"),0) = 0 Then
			Response.Write oblog.htm2js_div(""&digg&"","ob_logd"&id)
		End if
		Response.Write oblog.htm2js_div(""&Qs&"","ob_logm"&id)
	end If
	If OB_IIF(rs("isdigg"),1) = 1 And OB_IIF(rs("IsSpecial"),0) = 0 Then ShowDiggs (id)
	rs.close
	set rs=nothing
end sub

sub logs_count()
	dim i,strid,digg,Qs,diggstr
	id=oblog.filt_badstr(Trim(Request.QueryString("id")))
	if id="" then exit sub
	id=split(id,"$")
	for i=0 to Ubound(id)
		if id(i)<>"" then
			if strid="" then
				strid=CLng(id(i))
			else
				strid=strid&","&CLng(id(i))
			end if
		end if
	next
	set rs=oblog.execute("select logid,iis,commentnum,trackbacknum,diggnum,isDigg,IsSpecial from oblog_log a INNER JOIN oblog_user b ON a.authorid = b.userid where logid in ("&strid&")")
	while not rs.eof
	If OB_IIF(rs("IsSpecial"),0) = 0 Then
		digg = ShowDigg
		digg = Replace(digg,"$diggnum$",OB_IIF(rs(4),0))
		digg = Replace(digg,"$logid$",rs(0))
		digg = Replace(digg,"$showmsg$","推荐")
	End If
'		Qs = "|<ul class=""Reflect"" style=""display:inline;"" id=""menu_"&rs(0)&""" onmouseover=""menuFix('"&rs(0)&"');""> <li> <a href=""javascript:void(null)"">反映问题</a> <ul>"&GetMenuList(rs(0))&"</ul>"
		If IsArray(oblog.CacheReport)Then
			If UBound(oblog.CacheReport) > 0 Then
				Qs = "| <a href=""javascript:void(null)"" onclick=""openScript('"&blogurl&"report.asp?logid="&rs(0)&"',450,400)"">反映问题</a> "
			End If
		End If
		diggstr=diggstr& oblog.htm2js_div("("&rs(1)&")","ob_logr"&rs(0))
		diggstr=diggstr& oblog.htm2js_div("("&rs(2)&")","ob_logc"&rs(0))
		diggstr=diggstr& oblog.htm2js_div("("&rs(3)&")","ob_logt"&rs(0))
		If OB_IIF(rs("isDigg"),1) = 1  And OB_IIF(rs("IsSpecial"),0)=0  Then
			diggstr=diggstr& oblog.htm2js_div(""&digg&"","ob_logd"&rs(0))
		End if
		diggstr=diggstr& oblog.htm2js_div(""&Qs&"","ob_logm"&rs(0))
		rs.movenext
	wend
	set rs=Nothing
	response.write diggstr
end sub

sub comment_code()
	if oblog.cacheConfig(30)=1 then
		Response.Write(oblog.htm2js("验证码：<input name=""CodeStr"" type=""text"" size=""6"" maxlength=""20"" />"&oblog.getcode, False))
	end if
end sub

sub comment_code31()
	Dim tmpstr
	Randomize
	tmpstr=CStr(Int(900000*Rnd)+100000)
	if oblog.cacheConfig(30)=1 then
		'Response.Write(oblog.htm2js_div("验证码：<input name=""CodeStr"" type=""text"" size=""6"" maxlength=""20"" />"&oblog.getcode&" ","ob_code"))
		Response.Write("var addcode_f=false;function addcode(){if(!addcode_f){"&oblog.htm2js_div("验证码：<input name=""CodeStr"" type=""text"" size=""6"" maxlength=""20"" /> "&oblog.getcode&" ","ob_code")&"}addcode_f=true;}")
	else
		Response.Write("function addcode(){return true;}")
	end if
end Sub
'DIGG暂存于此
Sub ShowDiggs(logid)
	Dim RSDIGG,ShowList,reurl
	if not IsObject(conn) then link_database
	Set RSDIGG = Server.CreateObject ("ADODB.RecordSet")
	RSDIGG.open "SELECT TOP 45 user_Icon1,user_dir,user_folder,user_domain,user_domainroot,custom_domain,a.username,nickname,blogname FROM oblog_user a INNER JOIN oblog_digg b  ON a.userid = b.userid WHERE b.diggtype=-1 AND b.logid =  "&logid,CONN,1,1
	If Not RSDIGG.EOF Then
		ShowList = ""
		Do While Not RSDIGG.EOF
			if true_domain=1 then
				if RSDIGG("custom_domain")="" or isnull(RSDIGG("custom_domain")) then
					reurl="http://"&RSDIGG("user_domain")&"."&RSDIGG("user_domainroot")
				else
					reurl="http://"&RSDIGG("custom_domain")
				end If
			else
				reurl=blogdir&RSDIGG(1)&"/"&RSDIGG(2)
			end If
			ShowList = ShowList & vbcrlf &"				<ul class=""ShowList_UL"" style=""float:left;margin:6px 5px;padding:6px 2px 0px 2px;border:1px #eee solid;width:64px;height:90px;overflow:hidden;background:#f4f4f4 url(images/cierre_pie.gif) no-repeat left top;text-align:center;"">" & vbcrlf
			ShowList = ShowList &"					<li class=""ShowList_UserIco""><a href="""&reurl&""" title="""&RSDIGG("blogname")&"""><img src="""&ProIco(RSDIGG(0),1)&""" class=""ob_face"" style=""border:none;expression(this.width >48 && this.height < this.width ? 48: true); height: expression(this.height > 48 ? 48: true);""/></a></li>" & vbcrlf
			ShowList = ShowList &"					<li class=""ShowList_UserName"" style=""text-align:center;width:62px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;""><a href="""&reurl&""" title="""&RSDIGG("username")&""">"&RSDIGG("username")&"</a></li>" & vbcrlf
			ShowList = ShowList &"				</ul>" & vbcrlf
			RSDIGG.MoveNext
		Loop
		ShowList = vbcrlf & "<div id=""ArchivesShowList"" style=""clear:both;margin:0 0 20px 0;height:1%;"">" & vbcrlf & "	<strong>此文章被以下用户所推荐：</strong>" & vbcrlf & "	<tabel id=""ArchivesShowListTable"" style=""width:100%;border:0;"">" & vbcrlf & "		<tr>" & vbcrlf & "			<td>" & ShowList
		ShowList = ShowList & "			</td>" & vbcrlf & "		</tr>" & vbcrlf & "	</table>" & vbcrlf & "</div>" & vbcrlf
		ShowList = oblog.htm2js(ShowList,False)
		Response.Write "if (chkdiv('morelog')) {document.getElementById('morelog').innerHTML+='" & ShowList & "';}"
	End If
	Set RSDIGG = Nothing
End Sub
Function GetMenuList(logid)
	Dim MenuList,ii,Report
	Report = oblog.CacheReport
	For ii = 0 To UBound(Report)
		MenuList = MenuList &"<li><a href=""javascript:void(null)"" onclick=""report('"&logid&"','"&ii&"');"">"&Report(ii)&"</a></li>"
	Next
	GetMenuList = MenuList
End Function
ShowDigg = oblog.htm2js(ShowDigg,False)
Sub DiggCshow()
End Sub
%>
var ShowDigg ;
function diggit(logid){
	<%If true_domain = 1 Then %>
	// 此处必须配合0703号以后的 Oblog Rewrite 组件才能用 P_pass 并不是用户的真实密码
	var Ajax = new oAjax("/ajaxserver.asp?action=digglog&fromurl=",show_returnsave);
	<%
			Dim P_username,P_pass,P_true
			P_true="0"
			P_username=escape(Request.Cookies(cookies_name)("username"))
			P_pass=oblog.DecodeCookie(Request.Cookies(cookies_name)("password"))
			If P_username<>"" And P_pass<>"" Then P_true="1"
			If P_true="1"  Then
	%>
	var arrKey = new Array("logid","puser","ppass","ptrue","");
	var arrValue = new Array(logid,"<%=P_username%>","<%=P_pass%>","<%=P_true%>","");
			<%else%>
	var arrKey = new Array("logid","");
	var arrValue = new Array(logid,"");
			<%End If
			 %>
	Ajax.Post(arrKey,arrValue);
	<%Else%>
	var Ajax = new oAjax("<%=blogurl%>ajaxServer.asp?action=digglog&fromurl=",show_returnsave);
	var arrKey = new Array("logid","");
	var arrValue = new Array(logid,"");
	Ajax.Post(arrKey,arrValue);
	<%End If
	%>
}
function show_returnsave(arrobj){
	if (arrobj){
		if (arrobj.length == 4) {
			ShowDigg ='<%=ShowDigg%>';
			ShowDigg = ShowDigg.replace('$diggnum$',arrobj[3])
			ShowDigg = ShowDigg.replace('$logid$',arrobj[2])
			ShowDigg = ShowDigg.replace('$showmsg$',arrobj[0])
			document.getElementById("ob_logd"+arrobj[2]).innerHTML = ShowDigg;
			return false;
		}
		switch (arrobj[1]){
		case '1':
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.set('src',arrobj[1]);
			oDialog.event(arrobj[0],'');
			oDialog.button('dialogOk',"");
			break;
		case '2':
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.set('src',arrobj[1]);
			oDialog.event(arrobj[0],'');
			oDialog.button('dialogOk',"");
			break;
		}
		}
	}
<%

Sub hidden()'这些已经不用了 先注释掉%>
function menuFix(obj) {
    var sfEls = document.getElementById("menu_"+obj).getElementsByTagName("li");
    for (var i=0; i<sfEls.length; i++) {
			sfEls[i].onclick=function() {
			this.className+=(this.className.length>0? " ": "") + "sfhover";
        }
			sfEls[i].onmouseover=function() {
			this.className+=(this.className.length>0? " ": "") + "sfhover";
		}
			sfEls[i].onMouseDown=function() {
			this.className+=(this.className.length>0? " ": "") + "sfhover";
        }
			sfEls[i].onMouseUp=function() {
			this.className+=(this.className.length>0? " ": "") + "sfhover";
        }
			sfEls[i].onmouseout=function() {
			this.className=this.className.replace(new RegExp("( ?|^)sfhover\\b"),"");
        }
    }
}
function report(logid,report_type){
	var Ajax = new oAjax("<%=blogurl%>ajaxServer.asp?action=savereport",show_returnsave);
	var arrKey = new Array("logid","report_type");
	var arrValue = new Array(logid,report_type);
	Ajax.Post(arrKey,arrValue);
}
<%End Sub %>
