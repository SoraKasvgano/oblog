<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
dim mainsql,classid,strurl,rsmain,show_list,ustr
dim keyword,selecttype,isbest,bstr1,userid
classid=Request.QueryString("classid")
If classid<>"" Then classid=CLng(classid)
keyword=EncodeJP(oblog.filt_badstr(Request("keyword")))
keyword = HTMLEncode(keyword)
selecttype=oblog.filt_badstr(Request("selecttype"))
isbest=CLng (Request.QueryString("isbest"))
userid=CLng (Request.QueryString("userid"))
call sysshow()
If isbest = 1 Then
	G_P_Show =  Replace (G_P_Show,"$show_title_list$","精华日志列表--"&oblog.cacheConfig(2))
Else
	G_P_Show =  Replace (G_P_Show,"$show_title_list$","最新日志列表--"&oblog.cacheConfig(2))
End if
strurl="list.asp"
if classid<>0 then
	set rsmain=oblog.execute("select id from oblog_logclass where parentpath like '"&classid&",%' OR parentpath like '%,"&classid&"' OR parentpath like '%,"&classid&",%'")
	while not rsmain.eof
		ustr=ustr&","&rsmain(0)
		rsmain.movenext
	wend
	ustr=classid&ustr
	mainsql=" and a.classid in ("&ustr&")"
	'mainsql=" and oblog_log.classid="&classid
	strurl="list.asp?classid="&classid
end if
if keyword<>"" then
	select case selecttype
	case "topic"
		mainsql=" and topic like '%"&keyword&"%'"
		strurl="list.asp?keyword="&keyword&"&selecttype="&selecttype
	case "logtext"
		if oblog.cacheConfig(26)= "1" then
			mainsql=" and logtext like '%"&keyword&"%'"
			strurl="list.asp?keyword="&keyword&"&selecttype="&selecttype
		else
			oblog.adderrstr("当前系统已经关闭日志内容搜索。")
			oblog.showerr
		end if
	case "id"
		'Response.Write "不支持此种检索方式"
		'Response.End
		mainsql=" and (author like '%"&keyword&"%' or c.blogname like '%" & keyword &"%')"
		strurl="list.asp?keyword="&keyword&"&selecttype="&selecttype
	case "username"
		mainsql=" and c.username like '%"&keyword&"%'"
		strurl="list.asp?keyword="&keyword&"&selecttype="&selecttype
	case "nickname"
		mainsql=" and c.nickname like '%"&keyword&"%'"
		strurl="list.asp?keyword="&keyword&"&selecttype="&selecttype
	end select
end if
if isbest=1 then
	mainsql=mainsql&" and isbest=1"
	if strurl="list.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="精华日志"
end if
if userid>0 then
	mainsql=mainsql&" and a.userid="&userid
	if strurl="list.asp" then
		strurl=strurl&"?userid="&userid
	else
		strurl=strurl&"&userid="&userid
	end if
end if
call sub_showlist(mainsql,strurl)
G_P_Show=Replace(G_P_Show,"$show_list$",show_list)
Response.Write G_P_Show&oblog.site_bottom

sub sub_showlist(sql,strurl)
	dim topn
	dim msql
	G_P_PerMax=Int(oblog.CacheConfig(36))
	G_P_FileName=strurl
	if Request("page")<>"" then
    	G_P_This=cint(Request("page"))
	else
		G_P_This=1
	end if
	topn=oblog.CacheConfig(37)
	If classid<>"" Then
		msql="select top "&topn&" a.topic,a.author,a.addtime,a.commentnum,a.logid,b.classname,b.id,a.userid,logfile,a.isbest from oblog_log a,oblog_logclass b,oblog_user c where a.classid=b.id and a.userid=c.userid and a.isdel=0 and ishide=0 and passcheck=1 and isdraft=0 and a.blog_password=0 and (a.is_log_default_hidden=0 or a.is_log_default_hidden is null)"&sql
	Else
		msql="select top "&topn&" a.topic,a.author,a.addtime,a.commentnum,a.logid,'无分类' as classname,'0' as id ,a.userid,logfile,a.isbest from oblog_log a,oblog_user c where a.userid=c.userid and a.isdel=0 and ishide=0 and passcheck=1 and isdraft=0 and a.blog_password=0 and (a.is_log_default_hidden=0 or a.is_log_default_hidden is null)"&sql
	End If
	msql=msql&" order by a.logid desc"
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'Response.Write(msql)
	if not IsObject(conn) then link_database
	rsmain.Open msql,Conn,1,1
	show_list= vbcrlf & "<table width=""100%"" class=""List_table_top"">" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf
	if keyword="" then
		show_list=show_list&"当前位置：<a href=""index.asp"">首页</a>→<a href=list.asp>日志列表</a>→$className$(共调用 $allNum$ 篇日志)"
	else
		select case selecttype
		case "topic"
			show_list=show_list&"当前位置：<a href=""index.asp"">首页</a>→搜索日志标题关键字“"&keyword&"”"
		case "logtext"
			show_list=show_list&"当前位置：<a href=""index.asp"">首页</a>→搜索日志内容关键字“"&keyword&"”"
		case "id"
			show_list=show_list&"当前位置：<a href=""index.asp"">首页</a>→搜索博客名称关键字“"&keyword&"”"
		case else
			show_list=show_list&"当前位置：<a href=""index.asp"">首页</a>→搜索关键字“"&keyword&"”"
		end select
	end If
	Dim bstr
	bstr = Trim(Request.ServerVariables("query_string"))
	bstr = Replace(bstr,"&isbest=1","")
	bstr = Replace(bstr,"isbest=1","")
	if bstr<>"" then bstr="list.asp?"&bstr&"&isbest=1" else bstr="list.asp?isbest=1"
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "		<td align=""right"">" & vbcrlf
	show_list= show_list & "<a href="""&bstr&""">查看精华日志</a>" & vbcrlf
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf
	show_list= show_list & GetSysClasses
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "		<td align=""right"">" & vbcrlf
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "</table>" & vbcrlf
	show_list= show_list & "<hr id=""list_hr1"" />" & vbcrlf
	If bstr1 <>"" Then
		show_list =Replace(show_list,"$className$",bstr1)
	End if
  	if rsmain.eof and rsmain.bof Then
		show_list =Replace(show_list,"$allNum$",0)
		show_list =Replace(show_list,"$className$","")
'		show_list=show_list & "<br>共调用0篇日志<br>"
		'show_list=show_list&"</table>"
	Else
		if classid<>0 Then
'			show_list=show_list&"→"&rsmain(5)&""
			show_list =Replace(show_list,"$className$",rsmain(5))
		Else
			show_list =Replace(show_list,"$className$","全部分类")
		End if
    	G_P_AllRecords=rsmain.recordcount
'		show_list=show_list & "共调用" & G_P_AllRecords & " 篇日志<br>"
		show_list =Replace(show_list,"$allNum$",G_P_AllRecords)
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
        	show_list=show_list&oblog.showpage(false,true,"篇日志")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rsmain.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list&oblog.showpage(false,true,"篇日志")
        	else
	        	G_P_This=1
           		getlist()
           		show_list=show_list&oblog.showpage(false,true,"篇日志")
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim i,strtopic,userurl
	Dim arrayList
	ReDim arrayList(Int(oblog.CacheConfig(37))-1)
	show_list= show_list & "<table width=""100%"" id=""ListLog"" class=""List_table"">" & vbcrlf
	show_list= show_list & "	<thead>" & vbcrlf
	show_list= show_list & "		<tr>" & vbcrlf
	show_list= show_list & "			<th class=""t1"">日志标题</th>" & vbcrlf
	show_list= show_list & "			<th class=""t2"" width=""100"" align=""center"">作者</th>" & vbcrlf
	show_list= show_list & "			<th class=""t3"" width=""60"" align=""center"">日期</th>" & vbcrlf
	show_list= show_list & "			<th class=""t4"" width=""50"" align=""center"">评论</th>" & vbcrlf
	show_list= show_list & "			" & vbcrlf
	show_list= show_list & "		</tr>" & vbcrlf
	show_list= show_list & "	</thead>" & vbcrlf
	show_list= show_list & "	<tbody>" & vbcrlf
	i = 0
	do while not rsmain.eof
		arrayList(i) = rsmain("userid")
		If rsmain("isbest")=1 Then
			strtopic="<font color=red>" & oblog.filt_html(rsmain(0)) & "</font>"
		Else
			strtopic=oblog.filt_html(rsmain(0))
		End If
		if oblog.strLength(strtopic)>50 then
			strtopic=oblog.InterceptStr(strtopic,47)&"..."
		end If
		show_list=show_list&"		<tr>" & vbcrlf
		show_list=show_list&"			<td class=""t1""><a href="""&rsmain(8)&""" title="""&oblog.filt_html(rsmain(0))&""" target=""_blank"">"&strtopic&"</a></td>" & vbcrlf
		show_list=show_list&"			<td class=""t2"" width=""100"" align=""center""><a href=""go.asp?userid="&rsmain(7)&""" target=_blank><span name=""nickname_"&rsmain("userid")&""" id=""nickname_"&rsmain("userid")&""">"&rsmain("userid")&"</span></a></td>" & vbcrlf
		show_list=show_list&"			<td class=""t3"" width=""60""  align=""center"">"&Mid(FormatDateTime(rsmain(2),2),6)&"</td>" & vbcrlf
		show_list=show_list&"			<td class=""t4"" width=""50""  align=""center"">"&rsmain(3)&"</td>" & vbcrlf
		show_list=show_list&"		</tr>" & vbcrlf
		rsmain.movenext
		i=i+1
		if i>=G_P_PerMax then exit do
	loop
	show_list=show_list&"</table>"
	show_list = show_list & oblog.GetNickNameById (arrayList,i,G_P_This)
end Sub
Function GetSysClasses()
	Dim rst,sReturn
	Set rst=conn.Execute("select * From oblog_logclass Where idtype=0")
	If rst.Eof Then
		sReturn=""
	Else
		Do While Not rst.Eof
			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
			rst.Movenext
		Loop
		sReturn = "<option value=999>请选择</option><option value=0>所有分类</option>" & VBCRLF & sReturn
'		sReturn="<form name=photoform method=get>日志分类：<select name=classid onchange=""this.form.submit()"">" & VBCRLF & sReturn & "</select></form>"
			sReturn="<form name=photoform method=get>日志分类：<select name=classid onchange=""this.form.submit()"">" & VBCRLF & oblog.SelectedClassString(2,0,0) & "</select></form>"
End If
	rst.Close
	Set rst=Nothing
	GetSysClasses = sReturn
End Function
%>
