<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%

dim mainsql,strurl,rsmain,bstr1,isbest,show_list,classid,keyword
strurl="groups.asp"
isbest=CInt(Request.QueryString("isbest"))
classid=Request.QueryString("classid")
keyword=Trim(Request("keyword"))
If keyword<>"" Then keyword=oblog.filt_badstr(keyword)
call sysshow()
If isbest = 1 Then
	G_P_Show =  Replace (G_P_Show,"$show_title_list$","推荐"&oblog.CacheConfig(69)&"列表--"&oblog.cacheConfig(2)  )
Else
	G_P_Show =  Replace (G_P_Show,"$show_title_list$","最新"&oblog.CacheConfig(69)&"列表--"&oblog.cacheConfig(2)  )
End if
if isbest=1 then
	mainsql=mainsql&" and isbest=1"
	if strurl="groups.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="→推荐" &oblog.CacheConfig(69)
end if
if keyword<>"" then
	mainsql=mainsql & " And t_name like '%" & keyword &"%'"
	if strurl="groups.asp" then
		strurl=strurl&"?keyword="&keyword
	else
		strurl=strurl&"&keyword="&keyword
	end if
	bstr1="→搜索" &oblog.CacheConfig(69)
end if 

If IsNumeric(classid) Then
	classid=CLng(classid)
	If classid>0 Then
		mainsql= mainsql & " and Classid=" & classid & " "
	end if
	if strurl="groups.asp" then
		strurl=strurl&"?classid=" & classid
	end if
End If


call sub_showuserlist(mainsql,strurl)
G_P_Show=Replace(G_P_Show,"$show_list$",show_list)
Response.Write G_P_Show&oblog.site_bottom
sub sub_showuserlist(sql,strurl)
	dim topn,msql,bstr
	G_P_PerMax=CLng(oblog.CacheConfig(78))
	G_P_FileName=strurl
	if Request("page")<>"" then
    	G_P_This=cint(Request("page"))
	else
		G_P_This=1
	end if
	msql="select top "&CLng(oblog.CacheConfig(79))&" teamid,t_ico,t_name,intro from oblog_team Where iState=3 "&mainsql&" order by teamid desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'Response.Write(msql)
	rsmain.Open msql,Conn,1,1
	show_list= vbcrlf & "<table width=""100%"" class=""List_table_top"">" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf

	show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→" &oblog.CacheConfig(69)& "(共调用 $allNum$ 个" &oblog.CacheConfig(69)& ")" & vbcrlf
	bstr=Trim(Request.ServerVariables("query_string"))
	bstr = Replace(bstr,"&isbest=1","")
	bstr = Replace(bstr,"isbest=1","")
	if bstr<>"" then bstr="groups.asp?"&bstr&"&isbest=1" else bstr="groups.asp?isbest=1"
	show_list= show_list & bstr1 & "		</td>" & vbcrlf
	show_list= show_list & "		<td align=""right"">" & vbcrlf
	show_list= show_list & "<a href="""&bstr&""">查看推荐" &oblog.CacheConfig(69)& "</a>" & vbcrlf
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf
	show_list=show_list & GetSysClasses & vbcrlf
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "</table>" & vbcrlf
	show_list= show_list & "<hr />" & vbcrlf
  	if rsmain.eof and rsmain.bof Then
		show_list =Replace(show_list,"$allNum$",0)
		show_list=show_list & "共调用0个" &oblog.CacheConfig(69)& "<br>"
	else
    	G_P_AllRecords=rsmain.recordcount
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
        	show_list=show_list&oblog.showpage(false,true,"个" &oblog.CacheConfig(69)& "")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rsmain.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list&oblog.showpage(false,true,"个" &oblog.CacheConfig(69)& "")
        	else
	        	G_P_This=1
           		getlist()
           		show_list=show_list&oblog.showpage(false,true,"个" &oblog.CacheConfig(69)& "")
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim i,n
	dim title
	i=0
	show_list =Replace(show_list,"$allNum$",G_P_AllRecords)
	show_list=show_list&"<table width=""100%"" id=""ListGroups"" class=""List_table"">"& vbcrlf
	'ob_debug G_P_PerMax,1
	do while not rsmain.eof
		show_list=show_list&"<tr>"& vbcrlf
		for n=1 to 4
			i=i+1			
			if rsmain.eof then
				show_list=show_list&"<td width=""25%""></td>"& vbcrlf
			else
				title=oblog.CacheConfig(69)&"简介:"&oblog.filt_html(rsmain("intro"))
				show_list=show_list&"<td align=""center""> <a href='group.asp?gid="&rsmain("teamid")&"' title='"&title&"' target='_blank'><img src='"&OB_IIF(rsmain("t_ico"),"images/default_groupico.gif")&"' height='90' width='120' border='0' /></a><br /><a href='group.asp?gid="&rsmain("teamid")&"'>"&rsmain("t_name")&"</a></td>"& vbcrlf
			
				if not rsmain.eof then rsmain.movenext
			end If
			if i>=G_P_PerMax then exit do
		next
		show_list=show_list&"</tr>"& vbcrlf
		
	loop
	show_list=show_list&"</table>"
end sub

Function GetSysClasses()
	Dim rst,sReturn
	Set rst=conn.Execute("select * From oblog_logclass Where idtype=2")
	If rst.Eof Then
		sReturn=""
	Else
		Do While Not rst.Eof
			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
			rst.Movenext
		Loop
		sReturn = "<option value=0>请选择</option><option value=0>所有分类</option>" & VBCRLF & sReturn
		sReturn="<form name=photoform method=get>" &oblog.CacheConfig(69)& "分类：<select name=classid onchange=""this.form.submit()"">" & VBCRLF & sReturn & "</select></td><td align=""right"">请输入关键字<input type=""text"" size=10 name=""keyword"" value="&keyword&"><input type=""submit"" value=""查询""></form>"
	End If
	rst.Close
	Set rst=Nothing
	GetSysClasses = sReturn
End Function
%>