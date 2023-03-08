<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/inc_tags.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
Dim sSql,rst,iTagId,iUserId,sKeyword,i,sType,sAll
Dim sPage,sContent,sTitle,sForm,sErr
Dim oTagName,oNum,oLastUpdate,iReturn


sForm="<form name=""tagform"" method=""post"" action=""" & blogdir & P_TAGS_SYSURL & """ ><tr><td>" & VBCRLF
sForm = sForm & "&nbsp;&nbsp;" & P_TAGS_DESC & "关键字&nbsp;&nbsp;<input type=""text"" name=""keyword"" size=20 value=""" & sKeyword & """>" & VBCRLF
sForm = sForm & "<input type=""submit"" value=""搜索"">" & VBCRLF
sForm = sForm & "</td></tr></form>" & VBCRLF

sType = LCase(Trim(Request.Querystring("t")))
iTagId = Trim(Request.Querystring("tagid"))
iUserId = Trim(Request.Querystring("userid"))
sKeyword= Trim(Request("keyword"))
sAll=Trim(Request.Querystring)

If sAll & sKeyword="" Then sType="hottags"
if iTagId<>"" then
	iTagId=CLng(iTagId)
end if
if iUserId<>"" then
	iUserId=CLng(iUserId)
end if

Call link_database()

select Case sType
	Case "hottags"
		sTitle="最热门的100个" & P_TAGS_DESC
		sContent=Tags_Hottags()
	Case "cloud"
		sTitle=P_TAGS_DESC & "云图"
		sContent=Tags_SystemTags(1)
	Case "list"
		sTitle=P_TAGS_DESC & "列表"
		sContent=Tags_SystemTags(0)
	Case "user"
		sTitle=P_TAGS_DESC & "用户"
		sContent=GetUsersByTag(iTagId)
	Case Else
		If sKeyword="" Then
			If iTagId="" Then
				Call GoErr("必须指定" & P_TAGS_DESC & " ID")
			Else
				If Not IsNumeric(iTagId) Then
					Call GoErr("错误的" & P_TAGS_DESC & " ID")
				Else

					iReturn=CINT(Tags_TagName(iTagId,oTagName,oNum,oLastUpdate))
					If iReturn=-1 Then
						oTagName="--"
					End If
					sTitle=P_TAGS_DESC & "&nbsp;:&nbsp;<font color=red>" & oTagName & "</font>"
					iTagId=CLng(iTagId)
				End If
			End If
			If iUserId="" Then
				sContent = Tags_TagBlogs("",iTagId)
			Else
				If Not IsNumeric(iUserId) Then
					Call GoErr("错误的用户ID")
				Else
					iUserId=CLng(iUserId)
					iReturn=CINT(Tags_TagName(iTagId,oTagName,oNum,oLastUpdate))
					If iReturn=-1 Then
						oTagName="--"
					End If
					sTitle="&nbsp;:&nbsp;<font color=red>" & GetUserInfo(iUserId) & "</font>," & P_TAGS_DESC & "&nbsp;:&nbsp;<font color=red>" & oTagName & "</font>"
					sContent = Tags_TagBlogs(iUserId,iTagId)
				End If
			End If
			If iUserId & iTagId ="" Then
				sTitle=P_TAGS_DESC & "云图"
				sContent=Tags_SystemTags(1)
			End If
		Else
			sKeyword=ProtectSql(sKeyword)
			sTitle="包含<font color=red>" & sKeyword & "</font>的" & P_TAGS_DESC
			sContent=Tags_SearchTag(sKeyword)
		End If
End select

Function GoErr(byval sErrMsg)
	Response.Redirect "err.asp?message=" & sErrmsg
End Function

sPage = vbcrlf & "<table width=""100%"" class=""List_table_top"">" & vbcrlf
sPage = sPage & "	<tr>" & vbcrlf
sPage = sPage & "		<td>" & vbcrlf
sPage = sPage & sForm & vbcrlf
sPage = sPage & "		</td>" & vbcrlf
sPage = sPage & "		<td align=""right"">" & vbcrlf
sPage = sPage & "		</td>" & vbcrlf
sPage = sPage & "	</tr>" & vbcrlf
sPage = sPage & "	<tr>" & vbcrlf
sPage = sPage & "		<td>" & vbcrlf
sPage = sPage & "当前位置：<a href='index.asp'>首页</a>→" & sTitle & vbcrlf
sPage = sPage & "		<td align=""right"">" & vbcrlf
sPage = sPage & "<a href=""" & blogdir & P_TAGS_SYSURL & "?t=hottags"" >热门" & P_TAGS_DESC &"</a>　　"
sPage = sPage & "<a href=""" & blogdir & P_TAGS_SYSURL & "?t=cloud"">" & P_TAGS_DESC &"云图</a>" & vbcrlf
sPage = sPage & "		</td>" & vbcrlf
sPage = sPage & "	</tr>" & vbcrlf
sPage= sPage & "</table>" & vbcrlf
sPage= sPage & "<hr />" & vbcrlf
sPage = sPage & sContent & vbcrlf
call sysshow()
G_P_Show =  Replace (G_P_Show,"$show_title_list$", "TAG列表--"&oblog.cacheConfig(2) )
G_P_Show=Replace(G_P_Show,"$show_list$",sPage)
Response.Write G_P_Show & oblog.site_bottom
If IsObject(conn) Then
	conn.close
	Set conn=Nothing
End If
%>
