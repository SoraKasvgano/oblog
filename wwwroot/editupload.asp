<!--#include file="inc/inc_syssite.asp"-->
<%
Response.Expires=0
if not oblog.checkuserlogined() then
	Response.write("登录后才能操作!"):Response.End()
else
	if oblog.l_ulevel=6 then
		Response.write("您未通过管理员审核!"):Response.End()
	end if
end if
function getExt(sFile)'ffilter
	dim sExt,Item
	sExt=""
	for each Item In split(sFile,".")
		sExt = Item
	next
	getExt=sExt
end function

function writeFileSelections()
	dim nIndex,bFileFound,sColor,rs,bDisplay,sExt,sIcon,iSelected,ffilter,sql,file_size
	nIndex=0
	bFileFound=false
	ffilter=Trim(Request("ffilter"))
	select case ffilter
		case "media"
		sql="(file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm')"
		case "flash"
		sql="(file_ext='swf')"
		case "( file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' )"
		case else
		sql="(1=1)"

	end select
	If Not IsObject(conn) Then link_database
	set rs=conn.execute("select * from oblog_upfile where userid="&oblog.l_uid&" and "&sql&" order by fileid desc")

	Response.Write "<div style='overflow:auto;height:222px;width:100%;margin-top:3px;margin-bottom:2px;'>" & VbCrLf
	Response.Write "<table border=0 cellpadding=2 cellspacing=0 width=100% height=100% >" & VbCrLf
	'Response.Write "<tr><td colspan=4 width='100%'><b>"&getTxt("Files")&"</b></td></tr>"
	sColor = "#e7e7e7"
	while not rs.eof

		'ffilter ~~~~~~~~~~
		bDisplay=false
		sExt=getExt(rs("file_path"))
		file_size = rs("file_size")
		If file_size = "" Or IsNull(file_size) Then
			file_size = 0
		End if
		if ffilter="flash" then
			if LCase(sExt)="swf" then bDisplay=true
		elseif ffilter="media" then
			if LCase(sExt)="avi" or LCase(sExt)="wmv" or LCase(sExt)="mpg" or _
			   LCase(sExt)="mpeg" or LCase(sExt)="wav" or LCase(sExt)="wma" or _
			   LCase(sExt)="mid" or LCase(sExt)="mp3" then bDisplay=true
		elseif ffilter="image" then
			if LCase(sExt)="gif" or LCase(sExt)="jpg" or LCase(sExt)="png" then bDisplay=true
		else 'all
			bDisplay=true
		end if
		'~~~~~~~~~~~~~~~~~~

		if bDisplay then

			bFileFound=true

			nIndex=nIndex+1
			if sColor = "#EFEFF5" then
				sColor = ""
			else
				sColor = "#EFEFF5"
			end if

			'icons
			sIcon="ico_unknown.gif"
			If LCase(sExt)="asp" then sIcon="ico_asp.gif"
			If LCase(sExt)="bmp" then sIcon="ico_bmp.gif"
			If LCase(sExt)="css" then sIcon="ico_css.gif"
			If LCase(sExt)="doc" then sIcon="ico_doc.gif"
			If LCase(sExt)="exe" then sIcon="ico_exe.gif"
			If LCase(sExt)="gif" then sIcon="ico_gif.gif"
			If LCase(sExt)="htm" then sIcon="ico_htm.gif"
			If LCase(sExt)="html" then sIcon="ico_htm.gif"
			If LCase(sExt)="jpg" then sIcon="ico_jpg.gif"
			If LCase(sExt)="js"	 then sIcon="ico_js.gif"
			If LCase(sExt)="mdb" then sIcon="ico_mdb.gif"
			If LCase(sExt)="mov" then sIcon="ico_mov.gif"
			If LCase(sExt)="mp3" then sIcon="ico_mp3.gif"
			If LCase(sExt)="pdf" then sIcon="ico_pdf.gif"
			If LCase(sExt)="png" then sIcon="ico_png.gif"
			If LCase(sExt)="ppt" then sIcon="ico_ppt.gif"
			If LCase(sExt)="mid" then sIcon="ico_sound.gif"
			If LCase(sExt)="wav" then sIcon="ico_sound.gif"
			If LCase(sExt)="wma" then sIcon="ico_sound.gif"
			If LCase(sExt)="swf" then sIcon="ico_swf.gif"
			If LCase(sExt)="txt" then sIcon="ico_txt.gif"
			If LCase(sExt)="vbs" then sIcon="ico_vbs.gif"
			If LCase(sExt)="avi" then sIcon="ico_video.gif"
			If LCase(sExt)="wmv" then sIcon="ico_video.gif"
			If LCase(sExt)="mpeg" then sIcon="ico_video.gif"
			If LCase(sExt)="mpg" then sIcon="ico_video.gif"
			If LCase(sExt)="xls" then sIcon="ico_xls.gif"
			If LCase(sExt)="zip" then sIcon="ico_zip.gif"

			Response.Write "<tr style='background:" & sColor & "'>" & VbCrLf & _
				"<td><img src='editor/images/file/"&sIcon&"'></td>" & VbCrLf & _
				"<td valign=top width=100% ><u id=""idFile"&nIndex&""" style='cursor:pointer;' onclick=""selectFile('" & blogdir & rs("file_path") & "',1)"">" & rs("file_name") & "</u>&nbsp;&nbsp;<img style='cursor:pointer;' onclick=""downloadFile('" & blogdir & rs("file_path") & "')"" src='editor/images/download.gif'></td>" & VbCrLf & _
				"<td valign=top align=right nowrap>" & FormatNumber(file_size/1000,1) & " kb&nbsp;</td>" & VbCrLf & _
				"<td valign=top nowrap onclick=""deleteFile(" & nIndex & ")"">"
			Response.Write "</td></tr>" & VbCrLf
		end if
		rs.movenext
	wend
	if bFileFound=false then
		Response.Write "<tr><td colspan=4 height=100% align=center>无可选择的文件</td></tr></table></div>"
	else
		Response.Write "<tr><td colspan=4 height=100% ></td></tr></table></div>"
	end if

	Response.Write "<input type=hidden name=inpUploadedFile id=inpUploadedFile value='" & iSelected & "'>"
	Response.Write "<input type=hidden name=inpNumOfFiles id=inpNumOfFiles value='" & nIndex & "'>"
end function
%>

<base target="_self">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
body{font:8pt tahoma,arial,sans-serif;margin:0;background:#E9E8F2;color:#444444}
.dialogFooter{background-color:#E2E2ED;border-top:#CFCFCF 1px solid;}
td{font-size:8pt}
input {font:8pt tahoma,arial,sans-serif}
select {font:8pt tahoma,arial,sans-serif}
textarea {font:8pt tahoma,arial,sans-serif}
.inpSel {font:8pt tahoma,arial,sans-serif}
.inpTxt {font:8pt tahoma,arial,sans-serif;}
.inpChk {width:13;height:13;margin-right:3;margin-bottom:1}
.inpRdo {width:13;height:13;margin-right:3;margin-bottom:1}
.inpBtn {font:8pt tahoma,arial,sans-serif;}
.inpBtnOver {}
.inpBtnOut {}
</style>
<script>
/*Used for allocating problem:*/
<%If true_domain=1 Then%>
	var bReturnAbsolute=true;
<%Else%>
	var bReturnAbsolute=false;
<%End If%>
var activeModalWin;

function getAction(isUpload)//NEW 2.4
	{
	//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	//Clean previous ffilter=...
	sQueryString=window.location.search.substring(1)
	sQueryString=sQueryString.replace(/&upload=Y/,"")//NEW 2.4

	sQueryString=sQueryString.replace(/ffilter=media/,"")
	sQueryString=sQueryString.replace(/ffilter=image/,"")
	sQueryString=sQueryString.replace(/ffilter=flash/,"")
	sQueryString=sQueryString.replace(/ffilter=/,"")
	if(sQueryString.substring(sQueryString.length-1)=="&")
		sQueryString=sQueryString.substring(0,sQueryString.length-1)

	if(sQueryString.indexOf("=")==-1)
		{//no querystring
		sAction="editupload.asp?ffilter="+document.getElementById("selFilter").value;
		}
	else
		{
		sAction="editupload.asp?"+sQueryString+"&ffilter="+document.getElementById("selFilter").value
		}

	if(isUpload) sAction+="&upload=Y";//NEW 2.4
	//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	return sAction;
	}

function applyFilter()//ffilter
	{
	var Form1 = document.forms.Form1;

//	Form1.elements.inpCurrFolder.value=document.getElementById("selCurrFolder").value;
//	Form1.elements.inpFileToDelete.value="";
//
	Form1.action=getAction()
	Form1.submit()
	}
function upload(fpath)
	{
		selectFile(fpath,0);
	}
function modalDialogShow(url,width,height)//moz
    {
    var left = screen.availWidth/2 - width/2;
    var top = screen.availHeight/2 - height/2;
    activeModalWin = window.open(url, "", "width="+width+"px,height="+height+",left="+left+",top="+top);
    window.onfocus = function(){if (activeModalWin.closed == false){activeModalWin.focus();};};
    }
function downloadFile(fpath)
	{
	sFile_RelativePath = fpath;
	sFile_RelativePath = window.location.protocol + "//" + window.location.host.replace(/:80/,"")  + sFile_RelativePath
	window.open(sFile_RelativePath)
	}
function selectFile(fpath,sType)

	{
	sFile_RelativePath = fpath;

	//This will make an Absolute Path
	if(bReturnAbsolute)
		{
			var blogdir ;
			if (sType == 0)	{
				blogdir ='<%=blogdir%>'
			}
			else {
				blogdir = '/'
			}
			//sFile_RelativePath = window.location.protocol + "//" + window.location.host.replace(/:80/,"") + "/" + sFile_RelativePath;
			var thisurl;
			thisurl= window.location.host.replace(/:80/,"") + blogdir + sFile_RelativePath;
			thisurl=thisurl.replace("//","/");
			sFile_RelativePath = window.location.protocol + "//" + thisurl;
		//Ini input dr yg pernah pake port:
		//sFile_RelativePath = window.location.protocol + "//" + window.location.host.replace(/:80/,"") + "/" + sFile_RelativePath.replace(/\.\.\//g,"")
		}

	document.getElementById("inpSource").value=sFile_RelativePath;

	var arrTmp = sFile_RelativePath.split(".");
	var sFile_Extension = arrTmp[arrTmp.length-1]
	var sHTML="";

	//Image
	if(sFile_Extension.toUpperCase()=="GIF" || sFile_Extension.toUpperCase()=="JPG" || sFile_Extension.toUpperCase()=="PNG")
		{
		sHTML = "<img src=\"" + sFile_RelativePath + "\" >"
		}
	//SWF
	else if(sFile_Extension.toUpperCase()=="SWF")
		{
		sHTML = "<object "+
			"classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' " +
			"width='100%' "+
			"height='100%' " +
			"codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0'>"+
			"	<param name=movie value='"+sFile_RelativePath+"'>" +
			"	<param name=quality value='high'>" +
			"	<embed src='"+sFile_RelativePath+"' " +
			"		width='100%' " +
			"		height='100%' " +
			"		quality='high' " +
			"		pluginspage='http://www.macromedia.com/go/getflashplayer'>" +
			"	</embed>"+
			"</object>";
		}
	//Video
	else if(sFile_Extension.toUpperCase()=="WMV"||sFile_Extension.toUpperCase()=="AVI"||sFile_Extension.toUpperCase()=="MPG")
		{
		sHTML = "<embed src='"+sFile_RelativePath+"' hidden=false autostart='true' type='video/avi' loop='true'></embed>";
		}
	//Sound
	else if(sFile_Extension.toUpperCase()=="WMA"||sFile_Extension.toUpperCase()=="WAV"||sFile_Extension.toUpperCase()=="MID")
		{
		sHTML = "<embed src='"+sFile_RelativePath+"' hidden=false autostart='true' type='audio/wav' loop='true'></embed>";
		}
	//Files (Hyperlinks)
	else
		{
		sHTML = "<br><br><br><br><br><br>Not Available"
		}

	document.getElementById("idPreview").innerHTML = sHTML;
	}
bOk=false;
function doOk()
	{
	if(navigator.appName.indexOf('Microsoft')!=-1)
		window.returnValue=inpSource.value;
	else
		window.opener.setAssetValue(document.getElementById("inpSource").value);
	bOk=true;
	self.close();
	}
function doUnload()
	{
	if(navigator.appName.indexOf('Microsoft')!=-1)
		if(!bOk)window.returnValue="";
	else
		if(!bOk)window.opener.setAssetValue("");
	}
</script>
</head>
<body onUnload="doUnload()" onLoad="this.focus();" style="overflow:hidden;margin:0px;">

<table width="100%" height="100%" align=center style="" cellpadding=0 cellspacing=0 border=0 >
<tr>
<td valign=top style="background:url('bg.gif') no-repeat right bottom;padding-top:5px;padding-left:5px;padding-right:5px;padding-bottom:0px;">
		<table width=100% border="0">
		<tr>
		<td>
				<table cellpadding="2" cellspacing="2" border="0">
				<tr>
				<td valign=center nowrap></td>
				<td nowrap>
				</td>
				<td  width=100% align="right">
				<form name="Form1" id="Form1" action="" method="post"></form>
				<%
				'ffilter~~~~~~~~~
					dim sHTMLFilter,sAll,sMedia,sImage,sFlash,ffilter
					ffilter=Trim(Request("ffilter"))
					sHTMLFilter = "<select name=selFilter id=selFilter onchange='applyFilter()' class='inpSel'>"'ffilter
					sAll=""
					sMedia=""
					sImage=""
					sFlash=""
					if ffilter="" then sAll="selected"
					if ffilter="media" then sMedia="selected"
					if ffilter="image" then sImage="selected"
					if ffilter="flash" then sFlash="selected"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='' "&sAll&">所有文件</option>"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='media' "&sMedia&">媒体文件</option>"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='image' "&sImage&">图片文件</option>"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='flash' "&sFlash&">flash文件</option>"
					sHTMLFilter = sHTMLFilter & "</select>"
					Response.Write sHTMLFilter
				'~~~~~~~~~
				%>
				</td>
				</tr>
				</table>
		</td>
		</tr>
		<tr>
		<td valign=top align="center">

				<table width=100% cellpadding=0 cellspacing=0>
				<tr>
				<td>
					<div id="idPreview" style="text-align:center;overflow:auto;width:297;height:245;border:#d7d7d7 5px solid;border-bottom:#d7d7d7 3px solid;background:#ffffff;margin-right:2;"></div>
					<div align=center><input type="text" id="inpSource" name="inpSource" style="border:#cfcfcf 1px solid;width:295" class="inpTxt"></div>
				</td>
				<td valign=top width=100%>
					<%writeFileSelections()%>
				</td>
				</tr>
				</table>

		</td>
		</tr>
		<tr>
		<td>
		<iframe id='d_file' frameborder='0' src='upload.asp?tMode=10&re=' width='100%' height='60' scrolling='no'></iframe>
		</td>
		</tr>
		</table>

</td>
</tr>
<tr>
<td class="dialogFooter" style="height:40px;padding-right:15px;" align=right valign=middle>
	<table cellpadding=0 cellspacing=0 ID="Table2">
	<tr>
	<td>
	<input name="btnOk" id="btnOk" type="button" value=" 确定 " onClick="doOk()" class="inpBtn" onMouseOver="this.className='inpBtnOver';" onMouseOut="this.className='inpBtnOut'">
	</td>
	</tr>
	</table>
</td>
</tr>
</table>

</body>
</html>