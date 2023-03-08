<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
dim mainsql,usertype,strurl,rsmain,show_list
strurl="listblogstar.asp"
call sysshow()
G_P_Show =  Replace (G_P_Show,"$show_title_list$",  "博客之星--"&oblog.cacheConfig(2))
call sub_showuserlist(mainsql,strurl)
G_P_Show=Replace(G_P_Show,"$show_list$",show_list)
Response.Write G_P_Show&oblog.site_bottom
sub sub_showuserlist(sql,strurl)
	dim topn,msql
	G_P_PerMax=Int(oblog.CacheConfig(36))
	G_P_FileName=strurl
	if Request("page")<>"" then
    	G_P_This=cint(Request("page"))
	else
		G_P_This=1
	end if
	msql="select TOP 500 * from [oblog_blogstar] where ispass=1 order by addtime desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'Response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "共调用0位博客之星<br>"
	else
    	G_P_AllRecords=rsmain.recordcount
		'show_list=show_list & "共调用" & G_P_AllRecords & " 位博客之星<br>"
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
        	show_list=show_list&oblog.showpage(false,true,"位博客之星")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rsmain.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list&oblog.showpage(false,true,"位博客之星")
        	else
	        	G_P_This=1
           		getlist()
           		show_list=show_list&oblog.showpage(false,true,"位博客之星")
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim rstmp,i
	show_list= vbcrlf & "<table width=""100%"" class=""List_table_top"">" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf
	show_list= show_list & "当前位置：<a href=""index.asp"">首页</a>→所有博客之星(共有" & G_P_AllRecords & "位)"
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "</table>" & vbcrlf
	show_list= show_list & "<hr />" & vbcrlf
	show_list= show_list & "<table width=""100%"" id=""ListBlogStar"" class=""List_table"">" & vbcrlf
	show_list= show_list & "	<thead>" & vbcrlf
	show_list= show_list & "		<tr>" & vbcrlf
	show_list= show_list & "			<th  class=""t1"" width=""160"" align=""center"">图片</th>" & vbcrlf
	show_list= show_list & "			<th  class=""t2"" width=""180"">博客</th>" & vbcrlf
	show_list= show_list & "			<th  class=""t3"">简介</th>" & vbcrlf
	show_list= show_list & "			" & vbcrlf
	show_list= show_list & "		</tr>" & vbcrlf
	show_list= show_list & "	</thead>" & vbcrlf
	show_list= show_list & "	<tbody>" & vbcrlf
    do while not rsmain.eof
		show_list=show_list&"		<tr>" & vbcrlf
		show_list=show_list&"			<td class=""t1"" width=""160"" align=""center""><a href='"&rsmain("userurl")&"' target='_blank'><img src="""&rsmain("picurl")&""" border=""0"" width=""120"" height=""90""  alt='"&oblog.filt_html(rsmain("blogname"))&"' /></a></td>" & vbcrlf
		show_list=show_list&"			<td class=""t2""><a href='"&rsmain("userurl")&"' target='_blank'>"&oblog.filt_html(rsmain("blogname"))&"</a></td>" & vbcrlf
		show_list=show_list&"			<td class=""t3"">"&oblog.filt_html(rsmain("info"))&"</td>" & vbcrlf
		show_list=show_list&"		</tr>" & vbcrlf
		rsmain.movenext
		i=i+1
		if i>=G_P_PerMax then exit do
	Loop
	show_list=show_list&"	</tbody>" & vbcrlf
	show_list=show_list&"</table>" & vbcrlf
end sub
%>