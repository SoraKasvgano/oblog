<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%

dim mainsql,strurl,rsmain,bstr1,isbest,show_list,classid,strPlayerUrl
strurl="photo.asp"
isbest=cint(Request.QueryString("isbest"))
classid=Request.QueryString("classid")
call sysshow()
G_P_Show =  Replace (G_P_Show,"$show_title_list$","最新相册列表--"&oblog.cacheConfig(2)  )
if isbest=1 then
	mainsql=mainsql&" and user_isbest=1"
	if strurl="photo.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="→推荐相片"
end if

If IsNumeric(classid) Then
	classid=CLng(classid)
	If classid>0 Then
		mainsql= mainsql & " and sysClassid=" & classid & " "
	elseif classid=-1 then
		mainsql= mainsql & " and isBigHead=1 "
	end if
	if strurl="photo.asp" then
		strurl=strurl&"?classid=" & classid
	end if
End If

strPlayerUrl= Replace(strurl,"photo.asp","photoplayer.asp")
call sub_showuserlist(mainsql,strurl)
G_P_Show=Replace(G_P_Show,"$show_list$",show_list)
Response.Write G_P_Show&oblog.site_bottom
sub sub_showuserlist(sql,strurl)
	dim topn,msql
	G_P_PerMax=Int(oblog.CacheConfig(38))
	G_P_FileName=strurl
	if Request("page")<>"" then
    	G_P_This=cint(Request("page"))
	else
		G_P_This=1
	end if
	msql="select top "&CLng(oblog.CacheConfig(39))&" photo_path,photo_readme,userid,fileID,photo_Name from oblog_album b where 1=1 "&sql&" AND ( b.ishide = 0 OR b.ishide IS NULL  ) order by photoID desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
'	OB_DEBUG (msql),1
	rsmain.Open msql,Conn,1,1
	show_list= vbcrlf & "<table width=""100%"" class=""List_table_top"">" & vbcrlf
	show_list= show_list & "	<tr>" & vbcrlf
	show_list= show_list & "		<td>" & vbcrlf
	show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→相册(共调用 $allNum$ 个相片)" & bstr1 & vbcrlf
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "		<td align='right'>" & vbcrlf
	show_list= show_list & GetSysClasses & vbcrlf
	show_list= show_list & "		</td>" & vbcrlf
	show_list= show_list & "	</tr>" & vbcrlf
	show_list= show_list & "</table>" & vbcrlf
	show_list= show_list & "<hr />" & vbcrlf
  	if rsmain.eof and rsmain.bof Then
		show_list =Replace(show_list,"$allNum$",0)
		show_list=show_list & "共调用0个相片<br>"
	else
    	G_P_AllRecords=rsmain.recordcount
		'show_list=show_list & "共调用" & G_P_AllRecords & " 个相片<br>"
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
        	show_list=show_list&oblog.showpage(false,true,"个相片")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rsmain.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list&oblog.showpage(false,true,"个相片")
        	else
	        	G_P_This=1
           		getlist()
           		show_list=show_list&oblog.showpage(false,true,"个相片")
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim i,bstr,n,fso
	dim title,userurl,imgsrc,preImgSrc
	Dim arrayList,arrayUBound
	arrayUBound = Int(oblog.CacheConfig(38))-1
	'防止系统后台自定义每页显示的数目过少导致错误
	If arrayUBound < 3 Then arrayUBound = 3
	ReDim arrayList(arrayUBound)
	Set fso = Server.CreateObject(oblog.CacheCompont(1))
	show_list =Replace(show_list,"$allNum$",G_P_AllRecords)
	show_list=show_list&"<table width='100%'  align='center' cellpadding='0' cellspacing='1'>"& vbcrlf
	i = 0
	do while not rsmain.eof
		show_list=show_list&"<tr height='22'>"& vbcrlf
		for n=1 to 4
			if rsmain.eof then
				show_list=show_list&"<td width='25%'></td>"& vbcrlf
			Else
				arrayList(i) = rsmain("userid")
				'title="图片说明:"&oblog.filt_html(rsmain(1))
				'userurl="<a href='more.asp?id="& rsmain("logid") &"' target='_blank'>"
				userurl="<a href='go.asp?albumid="& rsmain("userid") &"' target='_blank'>"
				imgsrc=rsmain(0)

				preImgSrc=Replace(imgsrc,right(imgsrc,3),"jpg")
				preImgSrc=Replace(preImgSrc,right(preImgSrc,len(preImgSrc)-InstrRev(preImgSrc,"/")),"pre"&right(preImgSrc,len(preImgSrc)-InstrRev(preImgSrc,"/")))
				if Not fso.FileExists(Server.MapPath(preImgSrc)) then
					preImgSrc=imgsrc
				end If
				If oblog.CacheConfig(67) = "1" Then
					imgsrc = "attachment.asp?path="&imgsrc
				End If
'				if rsmain(5)<>"" then
'					userurl=userurl&oblog.filt_html(rsmain(5))&"</a>"
'				else
'					userurl=userurl&oblog.filt_html(rsmain(4))&"</a>"
'				end if
				userurl = userurl &"<span name=""nickname_"&rsmain("userid")&""" id=""nickname_"&rsmain("userid")&""">"&rsmain("userid")&"</span></a>"
				show_list=show_list&"<td align='center'> <a href='go.asp?albumid="& rsmain("userid") &"' title='"&title&"' target='_blank'><img src='"&preImgSrc&"' height='100' width='130' border='0' /></a><br />来自:"&userurl&"</td>"& vbcrlf
				i=i+1
				if not rsmain.eof then rsmain.movenext
			end If
			If I >=Int(oblog.CacheConfig(38)) Then Exit For
		next
		show_list=show_list&"</tr>"& vbcrlf
		if i>=G_P_PerMax then exit do

	loop
	show_list=show_list&"</table>"
	show_list = show_list & oblog.GetNickNameById (arrayList,i,G_P_This)
	set fso=nothing
end sub
'获取系统分类
Function GetSysClasses()
	Dim rst,sReturn
	Set rst=conn.Execute("select * From oblog_logclass Where idtype=1")
	If rst.Eof Then
		sReturn=""
	Else
		Do While Not rst.Eof
			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
			rst.Movenext
		Loop
		sReturn = "<option value=999>请选择</option><option value=0>所有分类</option>" & VBCRLF & sReturn
		sReturn =  sReturn&"<option value=-1>大头贴</option>"
		sReturn="<form name=photoform method=get>相册分类：<select name=classid onchange=""this.form.submit()"">" & VBCRLF & sReturn & "</select></form>"
	End If
	rst.Close
	Set rst=Nothing
	GetSysClasses = sReturn
End Function
%>