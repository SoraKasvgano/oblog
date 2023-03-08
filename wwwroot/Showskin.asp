<!--#include file="conn.asp"-->
<%
dim rs,skinid,userid,teamid,teamskinid
dim show,skinshowlog,i
skinid=CLng(Request("id"))
userid=CLng(Request("userid"))
teamid=CLng(Request("teamid"))
teamskinid=CLng(Request("teamskinid"))
if not IsObject(conn) then link_database
if skinid>0 then
	set rs=conn.execute("select skinmain,skinshowlog from oblog_userskin where id="&skinid)
	if rs.eof then
		Response.write "模板不存在！"
	else
		skinshowlog=rs(1)
		for i=1 to 6
			skinshowlog=skinshowlog+rs(1)
		next
		show=Replace(rs(0),"$show_log$",skinshowlog)
		Response.write show
	end if
elseif teamskinid>0 then
	set rs=conn.execute("select skinmain from oblog_teamskin where id="&teamskinid)
	if rs.eof then
		Response.write "模板不存在！"
	else
		Response.write rs(0)
	end if

elseif userid>0 then
	set rs=conn.execute("select bak_skin1,bak_skin2 from oblog_user where userid="&userid)
	if rs.eof then
		Response.write "用户不存在！"
	else
		if rs(0)="" or rs(1)="" or isnull(rs(0)) or isnull(rs(1)) then
			Response.Write("当前没有备份模板")
		else
			skinshowlog=rs(1)
			for i=1 to 6
				skinshowlog=skinshowlog+rs(1)
			next
			show=Replace(rs(0),"$show_log$",skinshowlog)
			Response.write show
		end if
	end if

elseif teamid>0 then
	set rs=conn.execute("select BAK_skin1,bak_skin2 from oblog_team where teamid="&teamid)
	if rs.eof then
		Response.write "用户不存在！"
	else
		if rs(0)=""  or isnull(rs(0)) Or rs(1)= "" Or IsNull(rs(1)) then
			Response.Write("当前没有备份模板")
		else
			Response.write RS(0)
		end if
	end if
end if
set rs=nothing
if isobject(conn) then conn.close:set conn=nothing
%>