<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->

<%
dim mainsql,classid,strurl,rsmain,show_list,ustr
dim keyword,selecttype,isbest,bstr1,userid
classid=request.QueryString("classid")
If classid<>"" Then classid=CLng(classid)
keyword=EncodeJP(oblog.filt_badstr(request("keyword")))
selecttype=oblog.filt_badstr(request("selecttype"))
isbest=CLng (request.QueryString("isbest"))
userid=CLng (request.QueryString("userid"))
call sysshow()
G_P_Show =  Replace (G_P_Show,"$show_title_list$","最新用户推荐日志列表--"&oblog.cacheConfig(2)  )
strurl="digg.asp"
if classid<>0 then
	set rsmain=oblog.execute("select id from oblog_logclass where parentpath like '"&classid&",%' OR parentpath like '%,"&classid&"' OR parentpath like '%,"&classid&",%'")
	while not rsmain.eof
		ustr=ustr&","&rsmain(0)
		rsmain.movenext
	wend
	ustr=classid&ustr
	mainsql=" and a.classid in ("&ustr&")"
	'mainsql=" and oblog_log.classid="&classid
	strurl="digg.asp?classid="&classid
end if
if keyword<>"" then
	select case selecttype
	case "topic"
		mainsql=" and topic like '%"&keyword&"%'"
		strurl="digg.asp?keyword="&keyword&"&selecttype="&selecttype
	case "logtext"
		if oblog.cacheConfig(26)=1 then
		mainsql=" and logtext like '%"&keyword&"%'"
		strurl="digg.asp?keyword="&keyword&"&selecttype="&selecttype
		else
		oblog.adderrstr("当前系统已经关闭日志内容搜索。")
		oblog.showerr
		end if
	case "id"
		mainsql=" and (author like '%"&keyword&"%' or c.blogname like '%" & keyword &"%')"
		strurl="digg.asp?keyword="&keyword&"&selecttype="&selecttype
	end select
end if
if userid>0 then
	mainsql=mainsql&" and a.userid="&userid
	if strurl="digg.asp" then
		strurl=strurl&"?userid="&userid
	else
		strurl=strurl&"&userid="&userid
	end if
end if
call sub_showlist(mainsql,strurl)
G_P_Show=replace(G_P_Show,"$show_list$",show_list)
response.Write G_P_Show&oblog.site_bottom

sub sub_showlist(sql,strurl)
	dim topn
	dim msql
'	G_P_PerMax=Int(oblog.CacheConfig(36))
	G_P_PerMax = 10
	G_P_FileName=strurl
	if request("page")<>"" then
    	G_P_This=cint(request("page"))
	else
		G_P_This=1
	end if
	topn=oblog.CacheConfig(37)
	If classid<>"" Then
		msql="SELECT a.*,b.classname FROM oblog_userdigg AS a INNER JOIN oblog_logclass AS b ON a.classid=b.id INNER JOIN oblog_log AS c ON a.logid = c.logid WHERE c.isdel=0 a.istate = 1 AND a.classid = "&classid
	Else
		msql="SELECT a.* FROM oblog_userdigg AS a INNER JOIN oblog_log AS c ON a.logid = c.logid WHERE a.istate = 1 AND c.isdel=0"
	End If
	msql=msql&" order by a.diggnum desc,diggid desc"
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'response.Write(msql)
	if not IsObject(conn) then link_database
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "共调用0篇日志<br>"
	else
    	G_P_AllRecords=rsmain.recordcount
		show_list=show_list & "共调用" & G_P_AllRecords & " 篇日志<br>"
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
	dim i,strtopic,userurl,bstr
	show_list= vbcrlf & "<table width=""100%"" class=""List_table_top"">" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf
	if keyword="" then
		if classid<>0 then
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→日志类别("&rsmain("classname")&")"
		else
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→用户推荐日志列表(所有类别)"
		end if
	else
		select case selecttype
		case "topic"
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索日志标题关键字“"&keyword&"”"
		case "logtext"
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索日志内容关键字“"&keyword&"”"
		case "id"
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索博客名称关键字“"&keyword&"”"
		case else
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索关键字“"&keyword&"”"
		end select
	end If
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "</table>" & vbcrlf
	show_list= show_list & "<hr />" & vbcrlf
	show_list= show_list & "<table width=""100%"" id=""ListDIGG"" class=""List_table"">" & vbcrlf
	show_list= show_list & "	<tbody>" & vbcrlf


	show_list=show_list&"<table class=""index_digg"" width=""100%"" border=""0"">" & vbcrlf
	do while not rsmain.eof
		strtopic=oblog.filt_html(rsmain("diggtitle"))
		if oblog.strLength(strtopic)>50 then
			strtopic=oblog.InterceptStr(strtopic,47)&"..."
		end If
		show_list=show_list&"	<tr>" & vbcrlf
		show_list=show_list&"		<td class=""digg_t1"" align=""center"">" & vbcrlf
		show_list=show_list&"			<div class=""digg_list"">" & vbcrlf
		show_list=show_list&"				<div class=""digg_number"" id=""log"&rsmain("logid")&""">"&rsmain("diggnum")&"</div>" & vbcrlf
		show_list=show_list&"				<div class=""digg_submit"" id=""log_img"&rsmain("logid")&"""><a href=""javascript:void(null)"" onclick=""diggit("&rsmain("logid")&");"">推荐</a></div>" & vbcrlf
		show_list=show_list&"			</div>" & vbcrlf
		show_list=show_list&"		</td>" & vbcrlf
		show_list=show_list&"		<td class=""digg_t2"">" & vbcrlf
		show_list=show_list&"			<div class=""digg_title""><a href='"&rsmain("diggurl")&"' title='"&oblog.filt_html(rsmain("diggtitle"))&"' target=_blank>"&strtopic&"</a></div>" & vbcrlf
		show_list=show_list&"			<div class=""digg_time""><a href=""go.asp?userid="&rsmain("authorid")&""" target=""_blank"">"&oblog.filt_html(rsmain("author"))&"</a>&nbsp;<span>submitted&nbsp;"&rsmain("addtime")&"</span></div>" & vbcrlf
		show_list=show_list&"			<div class=""digg_content"">"&rsmain("diggdes")&"</div>" & vbcrlf
		show_list=show_list&"		</td>" & vbcrlf
		show_list=show_list&"	</tr>" & vbcrlf
		rsmain.movenext
		i=i+1
		if i>=G_P_PerMax then exit do
	Loop
	show_list= show_list & "	</tbody>" & vbcrlf
	show_list=show_list&"</table>"
end sub
%>
<script>
function diggit(logid){
	var Ajax = new oAjax("ajaxServer.asp?action=digglog&fromurl=<%=Replace(oblog.GetUrl,"&","$")%>",show_returnsave);
	var arrKey = new Array("logid","");
	var arrValue = new Array(logid,"");
	Ajax.Post(arrKey,arrValue);
}
function show_returnsave(arrobj){
	if (arrobj){
			if (arrobj[3] !='')
			{
				document.getElementById("log"+arrobj[2]).innerHTML = arrobj[3];
			}
			document.getElementById("log_img"+arrobj[2]).innerHTML = arrobj[0];
		}
	}
</script>
