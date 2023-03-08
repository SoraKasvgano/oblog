<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
dim mainsql,usertype,strurl,rsmain,province,city,bstr1,isbest,show_list,ustr
strurl="listblogger.asp"
usertype=cint(Request.QueryString("usertype"))
isbest=cint(Request.QueryString("isbest"))
province=oblog.filt_badstr(Request("province"))
city=oblog.filt_badstr(Request("city"))
call sysshow()
G_P_Show =  Replace (G_P_Show,"$show_title_list$", "最新博客列表--"&oblog.cacheConfig(2) )
if usertype>0 then
	set rsmain=oblog.execute("select id from oblog_userclass where parentpath like '"&usertype&",%' OR parentpath like '%,"&usertype&"' OR parentpath like '%,"&usertype&",%'")
	while not rsmain.eof
		ustr=ustr&","&rsmain(0)
		rsmain.movenext
	wend
	ustr=usertype&ustr
	mainsql=" and oblog_user.user_classid in ("&ustr&")"
	strurl="listblogger.asp?usertype="&usertype
	'mainsql="and user_classid="&usertype
else
	mainsql=""
end if
if province<>"" then
	strurl=strurl&"?province="&province
	mainsql=mainsql&" and province='"&province&"'"
end if
if city<>"" then
	strurl=strurl&"&city="&city
	mainsql=mainsql&" and city='"&city&"'"
end if
if isbest=1 then
	mainsql=mainsql&" and user_isbest=1"
	if strurl="listblogger.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="→推荐博客"
end if

call sub_showuserlist(mainsql,strurl)
G_P_Show=Replace(G_P_Show,"$show_list$",show_list)
Response.Write G_P_Show&oblog.site_bottom
sub sub_showuserlist(sql,strurl)
	dim topn,msql
	G_P_PerMax=Int(oblog.CacheConfig(42))
	G_P_FileName=strurl
	if Request("page")<>"" then
    	G_P_This=cint(Request("page"))
	else
		G_P_This=1
	end if
	msql="select top "&oblog.CacheConfig(77)&" username,blogname,sex,useremail,qq,msn,log_count,homepage,adddate,userid,province,city from [oblog_user] where lockuser=0 "&sql&" and (is_log_default_hidden=0 or is_log_default_hidden is null) order by userid desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'Response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "共调用0位博客<br>"
	else
    	G_P_AllRecords=rsmain.recordcount
		'show_list=show_list & "共调用" & G_P_AllRecords & " 位博客<br>"
		if G_P_This<1 then
       		G_P_This=1
    	end if
    	if (G_P_This-1)*G_P_PerMax>G_P_AllRecords then
	   		if (G_P_AllRecords mod G_P_PerMax)=0 then
	     		G_P_This= G_P_AllRecords \ G_P_PerMax
		  	else
		      	G_P_This= G_P_AllRecords \ G_P_PerMax + 1
	   		end if

    	end if
	    if G_P_This=1 then
        	getlist()
        	show_list=show_list&oblog.showpage(false,true,"位博客")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rsmain.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list&oblog.showpage(false,true,"位博客")
        	else
	        	G_P_This=1
           		getlist()
           		show_list=show_list&oblog.showpage(false,true,"位博客")
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim rstmp,i,bstr
	dim title,userurl
	set rstmp=conn.execute("select classname from oblog_userclass where id="&usertype)
	show_list= vbcrlf & "<table width=""100%"" class=""List_table_top"">" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf
	if not rstmp.eof then
		show_list=show_list&"			当前位置：<a href=""index.asp"">首页</a>→博客类别("&rstmp(0)&")" & vbcrlf
	end if
	if usertype=0 then
		show_list=show_list&"			当前位置：<a href=""index.asp"">首页</a>→所有博客(共调用" & G_P_AllRecords & " 位博客)" & vbcrlf
	end if
	bstr=Trim(Request.ServerVariables("query_string"))
	if bstr<>"" then bstr="listblogger.asp?"&Replace(Replace(bstr,"&isbest=1",""),"isbest=1","")&"&isbest=1" else bstr="listblogger.asp?isbest=1"
	show_list= show_list & bstr1 & "		</td>" & vbcrlf
	show_list= show_list & "		<td align='right'>" & vbcrlf
	show_list= show_list & "			<a href='"&bstr&"'>查看推荐博客</a>" & vbcrlf
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "</table>" & vbcrlf
	set rstmp=Nothing
	show_list= show_list & "<hr />" & vbcrlf
	show_list= show_list & "<table width=""100%"" id=""ListBlogger"" class=""List_table"">" & vbcrlf
	show_list= show_list & "	<thead>" & vbcrlf
	show_list= show_list & "		<tr>" & vbcrlf
	show_list= show_list & "			<th class=""t1"" width=""100"" align=""center"">用户</th>" & vbcrlf
	show_list= show_list & "			<th class=""t2"">博客</th>" & vbcrlf
	show_list= show_list & "			<th class=""t3"" width=""120"" align=""center"">来自</th>" & vbcrlf
	show_list= show_list & "			<th class=""t4"" width=""50"" align=""center"">日志</th>" & vbcrlf
	show_list= show_list & "			" & vbcrlf
	show_list= show_list & "		</tr>" & vbcrlf
	show_list= show_list & "	</thead>" & vbcrlf
	show_list= show_list & "	<tbody>" & vbcrlf
     do while not rsmain.eof
			title="======== 用 户 信 息 ========" & vbcrlf & "性别："
		if rsmain(2)=1 then
			title=title& "男"
		else
			title=title& "女"
		end if
		title=title&vbcrlf & "QQ："
		if rsmain(4)<>"" then
			title=title& rsmain(4)
		else
			title=title& "未填"
		end if
		title=title& vbcrlf & "MSN："
		if rsmain(5)<>"" then
			title=title& rsmain(5)
		else
			title=title& "未填"
		end if
		title=title& vbcrlf & "主页："
		if rsmain(7)<>"" then
			title=title& rsmain(7)
		else
			title=title& "未填"
		end if
		title=title& vbcrlf & "注册：" & rsmain(8)
		show_list= show_list & "		<tr>" & vbcrlf
		show_list= show_list & "			<td class=""t1"" width=""100"" align=""center""><a href=""blog.asp?name=" &rsmain("username")&""" target=""_blank"">"&oblog.filt_html(rsmain(0))&"</td>" & vbcrlf
		show_list= show_list & "			<td class=""t2""><a href=""blog.asp?name=" &rsmain("username")&""" title=""点击进入"&oblog.filt_html(rsmain(0))&"的blog页面"& vbcrlf&title&""" target=""_blank"">"&oblog.filt_html(rsmain(1))&"</a></td>" & vbcrlf
		show_list= show_list & "			<td class=""t3"" width=""120"" align=""center""><a href=""listblogger.asp?province="&rsmain(10)&"&city="&rsmain(11)&""">"&rsmain(10)&rsmain(11)&"</a></td>" & vbcrlf
		show_list= show_list & "			<td class=""t4"" width=""50"" align=""center"">"&rsmain(6)&"</td>" & vbcrlf
		show_list= show_list & "		</tr>" & vbcrlf
		i=i+1
		if i>=G_P_PerMax then exit do
		rsmain.movenext
	Loop
	show_list= show_list & "	</tbody>" & vbcrlf
	show_list= show_list & "</table>"

end sub
%>