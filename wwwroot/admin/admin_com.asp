<!--#include file="inc/inc_sys.asp"-->
<%
Const C_Items=23
Dim action
action=Trim(Request("action"))
Dim aObjects(12)
aObjects(0) = "Scripting.FileSystemObject"
aObjects(1) = "adodb.connection"
'-----------------------
aObjects(2) = "JMail.Message"				'JMail 4.3
aObjects(3) = "CDONTS.NewMail"				'CDONTS
aObjects(4) = "Persits.MailSender"			'ASPEMAIL
'-----------------------
aObjects(5) = "Adodb.Stream"				'Adodb.Stream
aObjects(6) = "Persits.Upload"				'Aspupload3.0
aObjects(7) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
aObjects(8) = "DvFile.Upload"				'DvFile-Up V1.0
aObjects(9) = "LyfUpload.UploadFile"
'-----------------------
aObjects(10) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
aObjects(11) = "Persits.Jpeg"				'AspJpeg
aObjects(12) = "sjCatSoft.Thumbnail"		'sjCatSoft.Thumbnail V2.6

if action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if


Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

sub showconfig()
dim rs,ac
set rs=oblog.execute("select * From oblog_config Where id=2")
ac=Split(rs("ob_value"),"$$")
set rs=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>վ������</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�� վ �� �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="#SiteInfo">������</a> | <a href="#SiteOption">���ƶ���</a>
      | <a href="#user">�ϴ�(ͼƬ)���</a> | <a href="#show">�ʼ����</a>
    </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�� �� �� �� Ϣ</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<form method="POST" action="admin_com.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <td width="348" class="tdbg" height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td class="tdbg">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td width="348" class="tdbg" height=23>վ������·����<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td class="tdbg">���ݿ��ַ��</td>
  </tr>
  <tr>
    <td class="tdbg" height=23>FSO�ı���д��
      <%
      If IsObjInstalled(aObjects(0))=false Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
    <td class="tdbg">���ݿ�ʹ�ã�
      <%If Not IsObjInstalled(aObjects(1)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
  </tr>
  <tr>
    <td class="tdbg" height=23>Jmail���֧�֣�
      <%If Not IsObjInstalled(aObjects(2)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
    <td class="tdbg">CDONTS���֧�֣�
      <%If Not IsObjInstalled(aObjects(3)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
  </tr>
    <tr>
      <td height="25" colspan="2" class="topbg"><a name="SiteOption"></a><b>�����������</b></td>
    </tr>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >ϵͳ����ʹ��Scripting.FileSystemObject��Adodb.Stream���</td>
      <td height="25" > ������ķ�����������������������Ѿ�����,���޸����������</td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >Scripting.FileSystemObject��</td>
      <td height="25" ><%
      	If ac(1)="" Then
      		Call EchoInput("a1",40,100,"Scripting.FileSystemObject")
      	Else
      		Call EchoInput("a1",40,100,ac(1))
    	End If
      	%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >Adodb.Stream��</td>
      <td height="25" >
      	<%
      	If ac(1)="" Then
      		Call EchoInput("a2",40,100,"Adodb.Stream")
      	Else
      		Call EchoInput("a2",40,100,ac(2))
    	End If
      	%></td>
    </tr>
     <tr>
      <td height="25" colspan="2" class="topbg"><a name="show"></a><b>�ʼ�ѡ��</b></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25">˵��</td>
      <td height="25">���øò����󣬽����Կ����ʼ���Ч�Լ��Ȼ��ڡ������ʼ����ͶԷ�������Դ���ıȽϴ�������ѡ��
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�Ƿ������ʼ�ע����֤����</td>
      <td height="25" ><% Call EchoRadio("a3","","",ac(3))%></td>
    </tr>
<!--     <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >��ѡ����Ҫʹ�õ��ʼ�ѡ��</td>
      <td height="25" ><%=MakeSelect_Mail(ac(4))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" ><p>�ʼ���ַ</p></td>
      <td height="25" ><% Call EchoInput("a5",40,50,ac(5))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >SMTP����</td>
      <td height="25" ><% Call EchoInput("a6",40,50,ac(6))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" ><p>�û���</p></td>
      <td height="25" ><% Call EchoInput("a7",40,50,ac(7))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >����</td>
      <td height="25" ><% Call EchoInput("a8",40,50,ac(8))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�Ƿ�����SMTP��֤</td>
       <td height="25" ><% Call EchoRadio("a9","","",ac(9))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25" >�ʼ����Ͷ�����Ŀ(ÿ�����10��,ÿ��Ŀ���ַ��һ��)</td>
      <td height="25" ><% Call EchoInput("a10",40,50,ac(10))%></td>
    </tr> -->
    <tr>
      <td height="25" colspan="2" class="topbg"><a name="upload" id="user"></a><strong>�ϴ�ѡ��</strong></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" ><p>ѡȡ�ϴ������<br>
          (���Ե�<a href="http://www.oblog.cn" target="_blank">http://www.oblog.cn</a>����Aspupload3.0���)
        </p>
        </td>
      <td height="25" ><%=MakeSelect_Upload(ac(11))%>
      	</td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >ͼƬ����ͼ��ˮӡ���ÿ��أ�<br>
        (�������谲װAspJpeg������ɵ�<a href="http://www.oblog.cn" target="_blank">http://www.oblog.cn</a>����)
      </td>
      <td height="25" ><%=MakeSelect_Photo(ac(12))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ�ͼƬ���ˮӡ������Ϣ����Ϊ�ջ�0��:</td>
      <td height="25" ><% Call EchoInput("a13",40,50,ac(13))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ����ˮӡ�����С:</td>
      <td height="25" ><% Call EchoInput("a14",40,50,ac(14))%><b>px</b> </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ����ˮӡ������ɫ:</td>
      <td height="25" ><input type="text" name="a15" ID="d_a15" size=10 value="<%if ac(15)="" then Response.Write("#FFFFFF") else Response.Write(ac(15))%>">
        <img border=0 id="s_a15" src="images/rect.gif" style="cursor:pointer;background-Color:<%=ac(15)%>;" onclick="SelectColor('a15');" title="ѡȡ��ɫ!">
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ����ˮӡ��������:</td>
      <td height="25" ><select name="a16" id="a16">
          <option value="����" <%If ac(16)="����" Then Response.Write "selected"%>>����</option>
          <option value="����_GB2312" <%If ac(16)="����_GB2312" Then Response.Write "selected"%>>����</option>
          <OPTION value="Andale Mono" <%If ac(16)="Andale Mono" Then Response.Write "selected"%>>Andale Mono</OPTION>
          <OPTION value="Arial" <%If ac(16)="Arial" Then Response.Write "selected"%>>Arial</OPTION>
          <OPTION value="Arial Black" <%If ac(16)="Arial Black" Then Response.Write "selected"%>>Arial Black</OPTION>
          <OPTION value="Century Gothic"<%If ac(16)="Century Gothic" Then Response.Write "selected"%>>Century Gothic</OPTION>
          <OPTION value="Comic Sans MS" <%If ac(16)="Comic Sans MS" Then Response.Write "selected"%>>Comic Sans MS</OPTION>
          <OPTION value="Courier New" <%If ac(16)="Courier New" Then Response.Write "selected"%>>Courier New</OPTION>
          <OPTION value="Georgia" <%If ac(16)="Georgia" Then Response.Write "selected"%>>Georgia</OPTION>
          <OPTION value="Impact" <%If ac(16)="Impact" Then Response.Write "selected"%>>Impact</OPTION>
          <OPTION value="Tahoma" <%If ac(16)="Tahoma" Then Response.Write "selected"%>>Tahoma</OPTION>
          <OPTION value="Times New Roman" <%If ac(16)="Times New Roman" Then Response.Write "selected"%>>Times New Roman</OPTION>
          <OPTION value="Stencil" <%If ac(16)="Stencil" Then Response.Write "selected"%>>Stencil</OPTION>
          <OPTION value="Verdana" <%If ac(16)="Verdana" Then Response.Write "selected"%>>Verdana</OPTION>
        </select></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ�ˮӡ�����Ƿ����:</td>
      <td height="25" ><% Call EchoRadio("a17","","",ac(17))%>
        </select></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ�ͼƬ���ˮӡLOGOͼƬ��Ϣ����Ϊ�ջ�0��:<br>
        ��дLOGO��ͼƬ���·��
        <br/>
        ��Ը�Ŀ¼��·����������ڸ�Ŀ¼�£���ֱ��д·�� test.jpg
         <br/>
        ������ڸ�Ŀ¼�µ�imagesĿ¼�£���Ϊ images/test.jpg</td>
      <td height="25" ><% Call EchoInput("a18",40,50,ac(18))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ�ͼƬ���ˮӡ͸����:</td>
      <td height="25"><% Call EchoInput("a19",40,50,ac(19))%>��60%����д0.6 </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >ˮӡͼƬȥ����ɫ:<br>
        ����Ϊ����ˮӡͼƬ��ȥ����ɫ��</td>
      <td height="25" ><input type="text" name="a20" ID="d_a20" size=10 value="<%if ac(20)="" then Response.Write("#FFFFFF") else Response.Write(ac(20))%>">
        <img border=0 id="s_a20" src="images/rect.gif" style="cursor:pointer;background-Color:<%=ac(20)%>;" onclick="SelectColor('a20');" title="ѡȡ��ɫ!">
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >ˮӡ���ֻ�ͼƬ�ĳ���������:<br>
        ��ˮӡͼƬ�Ŀ�Ⱥ͸߶ȡ�</td>
      <td height="25" >�߶ȣ�<% Call EchoInput("a21",5,5,ac(21))%>����,
      	��ȣ�<% Call EchoInput("a22",5,5,ac(22))%>���� </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >�ϴ�ͼƬ���ˮӡLOGOλ������ :</td>
      <td height="25" ><select NAME="a23" id="upset_DrawXYType">
          <option value="0" <%if ac(23)=0 then Response.Write("selected")%>>����</option>
          <option value="1" <%if ac(23)=1 then Response.Write("selected")%>>����</option>
          <option value="2" <%if ac(23)=2 then Response.Write("selected")%>>����</option>
          <option value="3" <%if ac(23)=3 then Response.Write("selected")%>>����</option>
          <option value="4" <%if ac(23)=4 then Response.Write("selected")%>>����</option>
        </select></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >&nbsp;</td>
      <td height="25" >&nbsp; </td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg" > <input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " > </td>
    </tr>
  </table>

</form>
<script language="javascript">

function SelectColor(what){
	var dEL = document.all("d_"+what);
	var sEL = document.all("s_"+what);
	var arr = showModalDialog("../images/selcolor.html", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
	if (arr) {
		dEL.value=arr;
		sEL.style.backgroundColor=arr;
	}
}
</script>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>
<%

end sub

'�ϴ������Ⲣ����ѡ���
Function MakeSelect_Upload(sValue)
	Dim sRet,s1,s2,s3,s4
	select Case sValue
		Case "0"
			s2=" selected"
		Case "1"
			s3=" selected"
		Case "2"
			s4=" selected"
		Case Else
			s1=" selected"
	End select
	sRet="<select name=a11>" & vbcrlf
	sRet= sRet & "<option value=""999""" & s1 & ">�ر�</option>" & vbcrlf
	If IsObjInstalled(aObjects(5)) Then sRet= sRet & "<option value=""0""" & s2 & ">������ϴ�</option>" & vbcrlf
	If IsObjInstalled(aObjects(6)) Then sRet= sRet & "<option value=""1""" & s3 & ">Aspupload3.0��� </option>" & vbcrlf
	If IsObjInstalled(aObjects(7)) Then sRet= sRet & "<option value=""2""" & s4 & ">SA-FileUp 4.0���</option>" & vbcrlf
	sRet= sRet & "</select>"
	MakeSelect_Upload=sRet
	sRet=""
End Function

'�Զ��������ͼ��������ɱ�
Function MakeSelect_Photo(sValue)
	Dim sRet,sRet1,bRet,s1,s2,s3
	select Case sValue
		Case "0"
			s1=" selected"
		Case "1"
			s2=" selected"
		Case "2"
			s3=" selected"
		Case Else
			s1=" selected"
	End select
	If IsObjInstalled("Persits.Jpeg") Then
		bRet=true
		sRet1= "AspJpeg���<font color=red><b>��</b>������֧��!</font>"
	Else
		bRet=false
		sRet1= "AspJpeg���<b>��</b>��������֧��!"
	End If
	sRet= "<select name=a12>" & vbcrlf
	sRet= sRet & "<option value=""0""" & s1 & ">�ر�����ͼ��ˮӡЧ��</option>" & vbcrlf
	If bRet Then
		sRet= sRet & "<option value=""1""" & s2 & ">��������ͼ��ˮӡ����Ч��(�Ƽ�)</option>" & vbcrlf
		sRet= sRet & "<option value=""2""" & s3 & ">��������ͼ��ˮӡͼƬЧ��</option>" & vbcrlf
	End If
	sRet= sRet & "</select>&nbsp;&nbsp;(" & sRet1 & ")"
	MakeSelect_Photo=sRet
	sRet=""
	sRet1=""
End Function

'�Զ�����ʼ����������ѡ���
Function MakeSelect_Mail(sValue)
	Dim sRet,s1,s2,s3,s4
	select Case sValue
		Case "0"
			s2=" selected"
		Case "1"
			s3=" selected"
		Case "2"
			s4=" selected"
		Case Else
			s1=" selected"
	End select
	sRet="<select name=a4>" & vbcrlf
	sRet= sRet & "<option value=""999""" & s1 & ">�رջ�֧��</option>" & vbcrlf
	If IsObjInstalled(aObjects(2)) Then sRet= sRet & "<option value=""0""" & s2 & ">JMail���</option>" & vbcrlf
	If IsObjInstalled(aObjects(3)) Then sRet= sRet & "<option value=""1""" & s3 & ">CDONT(2000/2003�Դ�)</option>" & vbcrlf
	If IsObjInstalled(aObjects(4)) Then sRet= sRet & "<option value=""2""" & s4 & ">AspMail���</option>" & vbcrlf
	sRet= sRet & "</select>"
	MakeSelect_Mail=sRet
	sRet=""
End Function

Sub Saveconfig()
	If Request.QueryString <>"" Then Exit Sub
	Dim rs, i,sOpt
  	sOpt=""
  	For i=1 To C_Items
  		sOpt=sOpt & "$$" & Replace(Trim(Request.Form("a"&i)),"$","")
  	Next
  	sOpt=Now&sOpt
    If Not IsObject(conn) Then link_database
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open  "select * From oblog_config Where Id=2",conn,1,3
    If rs.Eof Then rs.AddNew
    rs("ob_value")=sOpt
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ReloadCache
	EventLog "�����޸ķ�����������õĲ���!",""
    Set oblog=Nothing
    Response.Redirect "admin_com.asp"
End Sub

%>