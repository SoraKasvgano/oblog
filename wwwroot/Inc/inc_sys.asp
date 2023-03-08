<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="../conn.asp"-->
<!--#include file="class_sys.asp"-->
<!--#include file="md5.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.start
chk_sysadmin
dim admin_name
sub chk_sysadmin()
	dim admin_password,sql,rs	
	admin_name=oblog.filt_badstr(session("adminname"))
	admin_password=oblog.filt_badstr(session("adminpassword"))
	if admin_name="" then
		Response.redirect "admin_login.asp"
		exit sub
	end if
	sql="select id from oblog_admin where username='" & admin_name & "' and password='"&admin_password&"'"
	set rs=oblog.execute(sql)
	if rs.bof and rs.eof then
		set rs=nothing
		Response.redirect "admin_login.asp"
		exit sub
	end if
	rs.close
	set rs=nothing		
end sub
%>
<SCRIPT LANGUAGE="JavaScript">
function CheckSel(Voption,Value)
{
	var obj = document.getElementById(Voption);
	for (i=0;i<obj.length;i++){
		if (obj.options[i].value==Value){
		obj.options[i].selected=true;
		break;
		}
	}
}
function chang_size(num,objname)
{
	var obj=document.getElementById(objname)
	if (parseInt(obj.rows)+num>=3) {
		obj.rows = parseInt(obj.rows) + num;	
	}
	if (num>0)
	{
		obj.width="90%";
	}
}
</script>
<style type="text/css">
div.showpage{
	CLEAR: both;
	text-align: center;
	width: 100%;
}
</style>