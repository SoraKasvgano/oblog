<!--#include file="inc/inc_sys.asp"-->
<%
'20070704:配置项目87项(85为验证模块判断,86为是否给加精或通过博星申请发送短信通知.87为自定义清理回收站天数)
Const C_Items=90
Dim Action
Action = Trim(Request("action"))
Select Case action
	Case "saveconfig" 
		Call Saveconfig
	Case "updateuserdomain"
		Call updateuserdomain()
	Case Else 
		Call Showconfig
End Select 

Sub updateuserdomain()
					Dim user_domainroot,Arr_domainroot,TEMP_domainroot
					TEMP_domainroot=Trim(oblog.CacheConfig(4))
					If InStr(TEMP_domainroot,"|")>0 Then
						Arr_domainroot=Split(TEMP_domainroot,"|")
						user_domainroot=Arr_domainroot(0)
					Else
						user_domainroot=TEMP_domainroot
					End If
oblog.execute("update oblog_user set user_domain=userid,user_domainroot='"&user_domainroot&"' where user_domain='' or user_domain is null")
oblog.ShowMsg "更新成功","close"
End Sub 

Sub Showconfig()
Dim rs,ac,sConfig,i
Set rs = oblog.execute("select ob_Value From oblog_config Where Id=1")
sConfig=rs(0)
ac=Split(sConfig,"$$")

'主动升级功能,务必需要更改常量C_Items
'序号从1开始
If UBound(ac)<C_Items Then
	For i=1 To (C_Items-UBound(ac))
		sConfig=sConfig & "$$0"
	Next
	'重新分割
	ac=Split(sConfig,"$$")
End If
Set oblog=Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>站点配置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">网站配置</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="#SiteInfo">网站信息配置</a> | <a href="#sysInfo">系统参数</a> |   <a href="#SiteOption">功能模块</a> | <a href="#sys">系统调用模块</a> |  <a href="#spam">垃圾防护模块</a> | <a href="#code">验证模块</a>  |  <a href="#biz">商业用户功能模块</a> | <a href="#reg">注册选项</a>  <br />| <a href="#log">日志选项</a> | <a href="#cmt">留言评论</a> | <a href="#group">圈子选项</a> </td>
    </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<br />
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">网站配置</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="admin_setup.asp" id="form1" name="form1" onsubmit="return CheckRadio();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>网站信息配置</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >网站名称<b>(纯文字,不支持Html)</b>：</td>
      <td  width="409" height="25"><% Call EchoInput("a1",40,50,ac(1))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >网站标题<b>(纯文字,不支持Html)</b>：</td>
      <td><% Call EchoInput("a2",40,50,ac(2))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >网站地址：<br>
        重要！请填写完整URL地址,如http://www.oblog.com.cn/,<font color="#FF0000">不能省略最后的/号</font>,此设置将影响到rss和trackback的正常运行。</td>
      <td><% Call EchoInput("a3",40,255,ac(3))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > 二级域名根：<br />请按照oblog.cn这样的形式书写，如有多个二级域名，请用&quot;|&quot;隔开，<font color="#FF0000">如关闭二级域名，请留空</font>：</td>
      <td><% Call EchoInput("a4",40,255,ac(4))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > 群组二级域名根：<br />请按照qq.oblog.cn这样的形式书写，如有多个二级域名，请用&quot;|&quot;隔开，不能和二级域名根重复.<font color="#FF0000">如关闭二级域名，请留空</font>：</td>
      <td><% Call EchoInput("a75",40,255,ac(75))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">是否开启二级域名用户连接：<br /><font color="#FF0000">如关闭二级域名，请选择否</font></td>
      <td><% Call EchoRadio("a5","","",ac(5))%>&nbsp;<font color="#FF0000">(如关闭或不支持二级域名，请选择否!如果您以前从未启用过此项请点这里<A HREF="admin_setup.asp?action=updateuserdomain" target="_blank">初始化用户二级选项</A>！)</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>站点关键字：<br />（更容易被搜索引擎找到,&quot;,&quot;号隔开）</td>
      <td><% Call EchoInput("a9",50,100,ac(9))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25">站点版权信息：<br />（显示在系统页面底部）：</td>
      <td><textarea name="a10" id='a10' cols="55" rows="5"><%=ac(10)%></textarea>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >站长信箱：</td>
      <td> <% Call EchoInput("a11",50,100,ac(11))%></td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="sysInfo" id="user"></a><strong>系统参数<font color="red">（请勿频繁更改）</font></strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >站内连接路径参数：<b>（默认相对路径）</b></td>
      <td><% Call EchoRadio("a55","绝对路径","相对路径",ac(55))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >上传目录：<b>（默认：UploadFiles）</b></td>
      <td><% Call EchoInput("a56",12,12,ob_iif(ac(56),"UploadFiles"))%><font color=red>（若指定其他目录，请手工建该目录,不推荐用用户目录做上传目录不便于安全设置）</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >广告目录：<b>（默认：GG）</b></td>
      <td><% Call EchoInput("a80",12,12,ob_iif(ac(80),"GG"))%><font color=red>（请手工建该目录，并确认对此目录有修改的权限）</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >日志保存目录：<b>（默认ARCHIVES目录）</b></td>
      <td><% Call EchoRadio("a57","用户根目录","ARCHIVES目录",ac(57))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >用户目录命名方式：<b>（默认用户名做目录）</b></td>
      <td><% Call EchoRadio("a58","用户ID","用户名",ac(58))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >服务器操作系统语言<b>（默认简体中文）</b>：</td>
      <td><% Call EchoRadio("a24","简体中文","其他",ac(24))%><font color=red>（如果服务器为简体中文不要选择其他，会降低生成效率）</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >服务器所在时区<b>（默认GMT+8.00）</b>：</td>
      <td>
	  <select name="a68" id="a68">
		<option value="">请设置您所在时区</option>
		<option value="-12">(GMT-12.00)国际日期变更线西</option>
		<option value="-11">(GMT-11.00)中途岛，萨摩亚群岛</option>
		<option value="-10">(GMT-10.00)夏威夷</option>
		<option value="-9">(GMT-9.00)阿拉斯加</option>
		<option value="-8">(GMT-8.00)太平洋时间（美国和加拿大）；蒂华纳</option>
		<option value="-7.a">(GMT-7.00)奇瓦瓦，拉巴斯，马扎特兰</option>
		<option value="-7.b">(GMT-7.00)山地时间（美国和加拿大）</option>
		<option value="-7.c">(GMT-7.00)亚利桑那</option>
		<option value="-6.a">(GMT-6.00)瓜达拉哈拉，墨西哥城，蒙特雷</option>
		<option value="-6.b">(GMT-6.00)萨斯喀彻温</option>
		<option value="-6.c">(GMT-6.00)中部时间（美国和加拿大）</option>
		<option value="-6.d">(GMT-6.00)中美洲</option>
		<option value="-5.a">(GMT-5.00)波哥大，利马，基多</option>
		<option value="-5.b">(GMT-5.00)东部时间（美国和加拿大）</option>
		<option value="-5.c">(GMT-5.00)印第安那州（东部）</option>
		<option value="-4.a">(GMT-4.00)大西洋时间（加拿大）</option>
		<option value="-4.b">(GMT-4.00)加拉加斯，拉巴斯</option>
		<option value="-4.c">(GMT-4.00)圣地亚哥</option>
		<option value="-3.a">(GMT-3.00)纽芬兰</option>
		<option value="-3.b">(GMT-3.00)巴西利亚</option>
		<option value="-3.c">(GMT-3.00)布宜诺斯艾利斯，乔治敦</option>
		<option value="-3.d">(GMT-3.00)格陵兰</option>
		<option value="-2">(GMT-2.00)中大西洋</option>
		<option value="-1.a">(GMT-1.00)佛得角群岛</option>
		<option value="-1.b">(GMT-1.00)亚速尔群岛</option>
		<option value="0">(GMT)格林威治标准时间，都柏林，爱丁堡，伦敦，里斯本</option>
		<option value="0.a">(GMT)卡萨布兰卡，蒙罗维亚</option>
		<option value="1.b">(GMT+1.00)阿姆斯特丹，柏林，伯尔尼，罗马，斯德哥尔摩，维也纳</option>
		<option value="1.c">(GMT+1.00)贝尔格莱德，布拉迪斯拉发，布达佩斯，卢布尔雅那，布拉格</option>
		<option value="1.d">(GMT+1.00)布鲁塞尔，哥本哈根，马德里，巴黎</option>
		<option value="1.e">(GMT+1.00)萨拉热窝，斯科普里，华沙，萨格勒布</option>
		<option value="1.f">(GMT+1.00)中非西部</option>
		<option value="2.a">(GMT+2.00)布加勒斯特</option>
		<option value="2.b">(GMT+2.00)哈拉雷，比勒陀利亚</option>
		<option value="2.c">(GMT+2.00)赫尔辛基，基辅，里加，索非亚，塔林，维尔纽斯</option>
		<option value="2.d">(GMT+2.00)开罗</option>
		<option value="2.e">(GMT+2.00)雅典，贝鲁特，伊斯坦布尔，明斯克</option>
		<option value="2.f">(GMT+2.00)耶路撒冷</option>
		<option value="3.a">(GMT+3.00)巴格达</option>
		<option value="3.b">(GMT+3.00)科威特，利雅得</option>
		<option value="3.c">(GMT+3.00)莫斯科，圣彼得堡，伏尔加格勒</option>
		<option value="3.d">(GMT+3.00)内罗毕</option>
		<option value="3.e">(GMT+3.00)德黑兰</option>
		<option value="4.a">(GMT+4.00)阿布扎比，马斯喀特</option>
		<option value="4.b">(GMT+4.00)巴库，第比利斯，埃里温</option>
		<option value="4.5">(GMT+4.30)喀布尔</option>
		<option value="5.a">(GMT+5.00)叶卡捷琳堡</option>
		<option value="5.b">(GMT+5.00)伊斯兰堡，卡拉奇，塔什干</option>
		<option value="5.5">(GMT+5.30)马德拉斯，加尔各答，孟买，新德里</option>
		<option value="5.75">(GMT+5.45)加德满都</option>
		<option value="6.a">(GMT+6.00)阿拉木图，新西伯利亚</option>
		<option value="6.b">(GMT+6.00)阿斯塔纳，达卡</option>
		<option value="6.c">(GMT+6.00)斯里哈亚华登尼普拉</option>
		<option value="6.d">(GMT+6.30)仰光</option>
		<option value="7.a">(GMT+7.00)克拉斯诺亚尔斯克</option>
		<option value="7.b">(GMT+7.00)曼谷，河内，雅加达</option>
		<option value="8.a">(GMT+8.00)北京，重庆，香港特别行政区，乌鲁木齐</option>
		<option value="8.b">(GMT+8.00)吉隆坡，新加坡</option>
		<option value="8.c">(GMT+8.00)珀斯</option>
		<option value="8.d">(GMT+8.00)台北</option>
		<option value="8.e">(GMT+8.00)伊尔库茨克，乌兰巴图</option>
		<option value="9.a">(GMT+9.00)大坂，东京，札幌</option>
		<option value="9.b">(GMT+9.00)汉城</option>
		<option value="9.c">(GMT+9.00)雅库茨克</option>
		<option value="9.501">(GMT+9.30)阿德莱德</option>
		<option value="9.502">(GMT+9.30)达尔文</option>
		<option value="10.a">(GMT+10.00)布里斯班</option>
		<option value="10.b">(GMT+10.00)符拉迪沃斯托克（海参崴）</option>
		<option value="10.c">(GMT+10.00)关岛，莫尔兹比港</option>
		<option value="10.d">(GMT+10.00)霍巴特</option>
		<option value="10.e">(GMT+10.00)堪塔拉，墨尔本，悉尼</option>
		<option value="11">(GMT+11.00)马加丹，索罗门群岛，新喀里多尼亚</option>
		<option value="12.a">(GMT+12.00)奥克兰，惠灵顿</option>
		<option value="12.b">(GMT+12.00)斐济，堪察加半岛，马绍尔群岛</option>
		<option value="13">(GMT+13.00)努库阿洛法</option>
	</select>
	</td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="SiteOption" id="user"></a><strong>功能模块</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>

    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">是否允许用户取回密码：</td>
      <td> <% Call EchoRadio("a84","","",ac(84))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否开启服务：</td>
      <td> <% Call EchoRadio("a12","","",ac(12))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">是否开启音乐盒：</td>
      <td> <% Call EchoRadio("a81","","",ac(81))%><font color="#FF0000">（必须先开启服务）</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">是否开启相册：</td>
      <td> <% Call EchoRadio("a76","","",ac(76))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否禁止接收引用通告<b>（永久禁止）</b>：</td>
      <td> <% Call EchoRadio("a54","","",ac(54))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">用户页面统计防刷新时间：</td>
      <td><% Call EchoInput("a31",10,10,Ob_IIF(ac(31),"30")) %>
        秒 </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">首页静态文件的更新时间：</td>
      <td><% Call EchoInput("a33",10,10,Ob_IIF(ac(33),"300"))%>秒 <font color="#FF0000">（建议设置300秒以上，否则将极耗费服务器资源）</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >转向用户首页是否转到INDEX文件：</td>
      <td> <% Call EchoRadio("a46","","",ac(46))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >编辑器是否启用XHTML查看源码：</td>
      <td> <% Call EchoRadio("a53","","",ac(53))%><font color="#FF0000">（不建议启用，文章字符过多容易造成浏览器假死）</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''" style="display:''">
      <td>自动读取短消息的时间（默认10分钟）：</td>
      <td><% Call EchoInput("a8",10,10,Ob_IIF(ac(8),"10"))%>分&nbsp;<font color="#FF0000">（数值不要太小，否则极耗费资源）</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >是否开启防盗链：</td>
      <td> <% Call EchoRadio("a67","","",ac(67))%>&nbsp;<font color="#FF0000">（不建议开启，耗费资源高）</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >是否允许游客浏览附件：</td>
      <td> <% Call EchoRadio("a82","","",ac(82))%>&nbsp;<font color="#FF0000">（建议不允许）</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >是否允许游客DIGG(推荐日志)：</td>
      <td> <% Call EchoRadio("a83","","",ac(83))%></td>
    </tr>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >是否自动发送短信通知给加精推荐或通过博星审核：</td>
      <td> <% Call EchoRadio("a86","","",ac(86))%>(对系统性能稍有影响)</td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="sys" id="user"></a><strong>系统调用模块</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">系统日志列表每页显示日志条数：</td>
      <td><% Call EchoInput("a36",10,10,Ob_IIF(ac(36),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25"><p>系统日志列表调用日志总条数：</p></td>
      <td><% Call EchoInput("a37",10,10,Ob_IIF(ac(37),"500"))%>（对应list.asp）</td>
	 </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">博客列表每页显示博客条数：</td>
      <td><% Call EchoInput("a42",10,10,Ob_IIF(ac(42),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">博客列表显示博客总条数：</td>
      <td><% Call EchoInput("a77",10,10,Ob_IIF(ac(77),"20"))%>（对应ListBlogger.asp）</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">系统相片列表每页显示相片个数：</td>
      <td><% Call EchoInput("a38",10,10,Ob_IIF(ac(38),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25"><p>系统相片列表调用相片总个数：</p></td>
      <td><% Call EchoInput("a39",10,10,Ob_IIF(ac(39),"500"))%>（对应photo.asp）</td>
	</tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">系统群组列表每页显示群组个数：</td>
      <td><% Call EchoInput("a78",10,10,Ob_IIF(ac(78),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25"><p>系统群组列表调用群组总个数：</p></td>
      <td><% Call EchoInput("a79",10,10,Ob_IIF(ac(79),"500"))%>（对应groups.asp）</td>
	</tr>
    <td height="25" class="topbg"><a name="spam" id="user"></a><strong>垃圾防护模块</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>引用通告授权码更新间隔时间：</p></td>
      <td> <% Call EchoInput("a64",10,10,Ob_IIF(ac(64),"120"))%>分钟（30分钟最小,1440分钟最大）</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>同一IP单位时间内允许的引用通告数目：<font color="#FF0000"><%=Chr(-23847)%></font></p></td>
      <td> <% Call EchoInput("a65",10,10,Ob_IIF(ac(65),"20"))%>条（超过则自动锁定IP）</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>检查引用通告数目的单位时间：</p></td>
      <td> <% Call EchoInput("a66",10,10,Ob_IIF(ac(66),"120"))%>分钟（控制编号一）</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>注册授权码更新间隔时间：</p></td>
      <td> <% Call EchoInput("a60",10,10,Ob_IIF(ac(60),"1440"))%>分钟（30分钟最小,1440分钟最大）</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>注册后多长时间可以发布日志：</p></td>
      <td> <% Call EchoInput("a19",10,10,Ob_IIF(ac(19),"20"))%>分钟</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>同一IP两次注册之间的间隔时间：</p></td>
      <td><% Call EchoInput("a20",10,10,Ob_IIF(ac(20),"300"))%>秒</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > 同一IP1小时内的注册限制数目：</td>
      <td> <% Call EchoInput("a21",10,10,Ob_IIF(ac(21),"20"))%>个（0为不限制，白名单内的IP除外）
        </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > 同一IP24小时内的注册限制数目：</td>
      <td> <% Call EchoInput("a14",10,10,Ob_IIF(ac(14),"50"))%>个（0为不限制，白名单内的IP除外）
        </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>同一IP单位时间内允许的评论、留言数目：<font color="#FF0000"><%=Chr(-23846)%></font></p></td>
      <td> <% Call EchoInput("a62",10,10,Ob_IIF(ac(62),"100"))%>条</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>单位时间内允许的评论、留言总数目：<font color="#FF0000"><%=Chr(-23845)%></font></p></td>
      <td> <% Call EchoInput("a63",10,10,Ob_IIF(ac(63),"100"))%>条</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>检查评论、留言数目的单位时间：</p></td>
      <td> <% Call EchoInput("a61",10,10,Ob_IIF(ac(61),"60"))%>分钟（控制编号二、三）</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >发生多少次敏感行为后系统自动封禁：</td>
      <td><% Call EchoInput("a13",10,10,Ob_IIF(ac(13),"5"))%>次（0为不限制）
    </tr>
	<tr>
      <td height="25" class="topbg"><a name="code" id="user"></a><strong>验证模块</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
		   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">是否开启注册邮件激活：</td>
      <td> <% Call EchoRadio("a88","","",Ob_IIF(ac(88),"0"))%>（需要邮件组件支持，并必须设置好相应的邮件服务器参数）</td>
    </tr>
			   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">是否默认屏蔽新用户的日志显示在系统首页：</td>
      <td> <% Call EchoRadio("a89","","",Ob_IIF(ac(89),"0"))%>（如果选是，您必须手动解除用户的前台文章显示屏蔽。）</td>
    </tr>
	  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td>验证模块是类型设置：</td>
      <td><input type="radio" name="a85" id="a85" value="0" <%If ac(85)=0 Then %>checked <%End If %> />只为数字验证码<input type="radio" name="a85" id="a85" value="1"  <%If ac(85)=1 Then %>checked <%End If %>  />只用自定义问题验证<input type="radio" name="a85" id="a85" value="2"  <%If ac(85)=2 Then %>checked <%End If %>  />混合验证方式.<br/><font color="red">(选择新版问题和混合验证方式的话登录页面默认为不使用问题验证)</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td>用户注册是否需要开启验证模块：</td>
      <td><% Call EchoRadio("a16","","",ac(16))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">用户登录是否需要开启验证模块：</td>
      <td><% Call EchoRadio("a29","","",ac(29))%>(登录在开启验证下验证模块默认为验证码)</td>
    </tr>
   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">用户发表评论，留言是否需要开启验证模块：</td>
      <td><% Call EchoRadio("a30","","",ac(30))%></td>
    </tr>
      <td height="25" class="topbg"><a name="biz" id="user"></a><strong>商业用户功能模块</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
     <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" ><a href="http://news.oblog.cn/news/20060110192.shtml" target="_blank" title="查看介绍">是否启用移动组件：</a></td>
      <td> <% Call EchoRadio("a51","","",ac(51))%>(通过彩信和邮件发布日志，<a href=" http://www.oblog.cn/gmzn.shtml" target="_blank" title="查看介绍"><font color=red>商业版本功能</font></a>)</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >启用移动组件后,用于接收数据的电子邮件地址(具体请咨询Oblog客服人员)：</td>
       <td><% Call EchoInput("a52",30,50,ac(52))%>(空间足够大，且不要启用太多过滤规则，防止接收不到数据)</td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="reg" id="user"></a><strong>注册选项</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否允许新用户：</td>
      <td> <% Call EchoRadio("a15","","",ac(15))%><font color="#FF0000">（关闭后，将只允许后台添加）</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >用户注册后是否自动创建目录（默认是）：</td>
      <td> <% Call EchoRadio("a59","","",ac(59))%><font color="#FF0000">（选否可节省磁盘空间，但可能造成不好的用户体验）</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许中文用户名<font color="#FF0000"></font>：</td>
      <td><% Call EchoRadio("a6","","",ac(6))%>&nbsp;<font color="#FF0000">（如果启用用户目录强制为userid）</font></td>
    </tr>
     <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>是否启用邀请码机制：<a href="#h2" onClick="hookDiv('hh2','')"><img src="images/ico_help.gif" border=0></a></p></td>
      <td> <% Call EchoRadio("a17","","",ac(17))%></td>
    </tr>
    <tr id="hh2" style="display:none" name="h2">
      <td colspan=2> <p>什么是邀请码</p>
        现有会员每天可以获取一定数量的邀请码，他可以将该邀请码手工发送给他人用于本站的注册<br/>
        (根据会员组的不同可用邀请码也不同,每个邀请码只能使用一次，且不可累积)<br/>
        新会员注册时，必须输入一个有效的邀请码才能进行注册，否则不允许注册。<br/>
        使用邀请码机制后，建议不要再启用注册审核，否则会因为注册步骤繁琐而给用户带来不好的体验
        </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>新用户注册是否需要管理员认证：<br>
          </p></td>
      <td> <% Call EchoRadio("a18","","",ac(18))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>邮件地址唯一：</p></td>
      <td> <% Call EchoRadio("a22","","",ac(22))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>昵称不允许重复：</p></td>
      <td> <% Call EchoRadio("a47","","",ac(47))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>博客名称不允许重复：</p></td>
      <td> <% Call EchoRadio("a48","","",ac(48))%></td>
    </tr>
      <td height="25" class="topbg"><a name="log"></a><strong>日志选项</strong></td>
      <td class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''" style="display:''">
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">单篇日志允许最多字数：</td>
      <td><% Call EchoInput("a34",10,10,Ob_IIF(ac(34),"50000"))%>
        字(英文字符) </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">单篇日志允许最多TAG数：</td>
      <td><% Call EchoInput("a73",10,10,Ob_IIF(ac(73),"10"))%>
         </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">单篇日志允许最多引用通告数：</td>
      <td><% Call EchoInput("a74",10,10,Ob_IIF(ac(74),"20"))%>
         </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">发表日志，系统日志分类是否为必须：</td>
      <td><% Call EchoRadio("a25","","",ac(25))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">是否允许日志全文搜索(建议关闭)：</td>
      <td> <% Call EchoRadio("a26","","",ac(26))%></td>
    </tr>
      <td width="348" height="25">日志自动保存为草稿的时间（默认2分钟）：</td>
      <td><% Call EchoInput("a7",10,10,Ob_IIF(ac(7),"2"))%>分&nbsp;<font color="#FF0000">（数值不要太小，否则极耗费资源）</font></td>
    </tr>
	 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >系统自动清理多少天以前的回收站日志：<br>
        </td>
      <td> <% Call EchoInput("a87",10,10,Ob_IIF(ac(87),"100"))%>天  （建议设置三个月以上 即 100 系统默认最小为 60 天）</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">全站更新每更新<%=P_BLOG_UPDATEPAUSE%>篇日志暂停的时间：</td>
      <td><% Call EchoInput("a28",10,10,Ob_IIF(ac(28),"5"))%>秒&nbsp;（0为不暂停，最大为60，一般设置为10即可）</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">页面数据载入时显示字符：</td>
      <td><% Call EchoInput("a41",30,50,Ob_IIF(ac(41),"载入中。。。"))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">图片自动缩小宽度（为零不缩放）：</td>
      <td><% Call EchoInput("a43",10,10,Ob_IIF(ac(43),"0"))%>像素</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >图片是否随鼠标滚轮缩放：</td>
      <td> <% Call EchoRadio("a44","","",ac(44))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >部分显示日志是否使用htm标记强化过滤：<br>
        </td>
      <td> <% Call EchoRadio("a45","","",ac(45))%>（若选择此项，所有除图片以外的标记都将被过滤掉）</td>
    </tr>
	 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >日志文件名为空时默认为：<br></td>
      <td><% Call EchoRadio("a23","日志ID自动编号","日志发表时间",ac(23))%></td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="cmt"></a><strong>评论与留言</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>

    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">是否允许游客发表评论及留言：</td>
      <td> <% Call EchoRadio("a27","","",ac(27))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">是否允许随机填充访客姓名：</td>
      <td> <% Call EchoRadio("a90","","",ac(90))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >留言与回复是否默认通过审核：</td>
      <td> <% Call EchoRadio("a50","","",ac(50))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">评论排序方式：</td>
      <td><% Call EchoRadio("a40","倒序","正序",ac(40))%>只对新注册和未设置排序方式的用户有效 
</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">评论及留言允许最多字数：</td>
      <td><% Call EchoInput("a35",10,10,Ob_IIF(ac(35),"2000"))%>
        字(英文字符) </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">留言，评论的时间间隔：</td>
      <td><% Call EchoInput("a32",10,10,Ob_IIF(ac(32),"60"))%>秒 </td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="group"></a><strong>圈子选项</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >申请圈子是否需要审核：</td>
      <td> <% Call EchoRadio("a49","","",ac(49))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">圈子名：</td>
      <td><% Call EchoInput("a69",10,10,Ob_IIF(ac(69),"群组"))%>&nbsp;<font color="#FF0000">（建议设置为两个字，如圈子，群组等）</font> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">圈子管理者名：</td>
      <td><% Call EchoInput("a70",10,10,Ob_IIF(ac(70),"群主"))%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">圈子人数上限：</td>
      <td><% Call EchoInput("a71",10,10,Ob_IIF(ac(71),"200"))%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">日志同时发布到圈子数目上限：</td>
      <td><% Call EchoInput("a72",10,10,Ob_IIF(ac(72),"3"))%></td>
    </tr>
    <tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg"> <a name="formbottom"></a><input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " > </td>
    </tr>
  </table>
</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>
<script language="JavaScript">
//定位时区的选项
var TimeZoneObject
TimeZoneObject = document.getElementsByTagName('option');
for (var i = 0;i < TimeZoneObject.length ; i ++ ){
	if (TimeZoneObject[i].value =='<%=ac(68)%>'){
		TimeZoneObject[i].selected=true;
	}
}
</script>
<%
Set rs = Nothing
End Sub

Sub Saveconfig()
	If Request.QueryString <>"" Then Exit Sub
	Dim rs, i,sOpt
	Dim arrayList
	ReDim arrayList(C_Items)
  	For i=1 To C_Items
  		sOpt=sOpt & "$$" & Replace(Trim(Request.Form("a"&i)),"$","")
		arrayList(i) = Replace(Trim(Request.Form("a"&i)),"$","")
  	Next
	Dim arrayDir
	arrayDir = Oblog.SysDir
	For i = 0 To UBound(arrayDir)
		if LCase(arrayList(56)) = arrayDir(i) Then
			oblog.ShowMsg "请勿选用系统目录作为上传目录",""
		End If
	Next
	On Error Resume Next
	'判断目录是否存在，如果不存在则自动创建
	Dim oFso
	Set oFso=Server.CreateObject(oblog.CacheCompont(1))
	If oFso.FolderExists(Server.Mappath(blogdir & LCase(arrayList(80)))) =False Then
		oFso.CreateFolder(Server.Mappath(blogdir & LCase(arrayList(80))))
	End If
	Set oFso=Nothing
	If Err Then
		Err.Clear
		oblog.ShowMsg "广告目录创建失败，请手工创建",""
	End if
  	sOpt=Now&sOpt
    If Not IsObject(conn) Then link_database
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open  "select * From oblog_config Where Id=1",conn,1,3
    If rs.Eof Then rs.AddNew
    rs("ob_value")=sOpt
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ReloadCache
	EventLog "进行修改网站信息配置的操作!",""
    Set oblog=Nothing
    Response.Redirect "admin_setup.asp"
End Sub
%>
<script language="javascript">
function CheckRadio()
{
	var obj = document.getElementsByTagName("input");
	for (var i = 0;i<obj.length ;i++ ){
		var e = obj[i];
		if (e.type == 'radio'){
			if (e.value !=1 &&e.value!=0 &&e.value!=2){
				alert('请确认每对单选按钮至少选中了一个选项');
				return false;
			}
		}
	}
}
</script>