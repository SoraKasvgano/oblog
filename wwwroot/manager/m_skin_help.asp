<html>
<head>
<title>oBlog 4.x 标签调用说明</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/admin/style.css" type="text/css" />
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">

<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">oBlog模板标记说明</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 colspan="2" class="topbg"><strong>系统模板标记说明</strong>
  <tr>
<td height=23 colspan="2" class="tdbg"><p><strong><font color=#0000ff>主模板（即系统首页，只对index.asp文件有效）：</font><br>
        注意</strong>：请使用正确的标签参数，否则可能出现不可预知的错误（一般为提示为下标越界）<br>
        　　　<font color=red>红色部分为4.0版本新增或与3.x版本不同参数的标签。</font></p></td>
  </tr>
  <tr>
    <td width="25%" height=23 class="tdbg"><p>$show_log(参数1,参数2,参数3,参数4,参数5,参数6,参数7,参数8,参数9)$<br>
      </p></td>
    <td width="75%" class="tdbg">此标记调用日志标题列表等信息。参数说明如下：<br>
      　　参数1：调用日志条数。<br>
      　　参数2：日志标题长度，以字符为单位，超过部分显示“...”。<br>
      　　参数3：排序方法，为1按发表时间，为2按点击数，为3按回复数。<br>
      　　参数4：是否精华，为1调用所有日志，为2调用精华日志。<br>
      　　参数5：调用多少天内的日志，以天为单位。<br>
      　　参数6：日志分类id，为0则调用所有分类的日志。<br>
      　　参数7：是否显示日志系统分类名，为1显示，为0不显示。<br>
      　　参数8：是否显示日志专题名，为1显示，为0不显示。<br>
      　　参数9：显示信息，1为显示发表时间和用户，2为显示发表时间，3为显示发表用户，4为显示发表用户和点击数，5为显示点击数，6为显示发表日期和用户，7为显示发表日期，8为显示回复数，0为不显示。</td>
  </tr>
  <tr>
<td width="25%" height=23 class="tdbg"><p>$show_userlog(参数1,参数2,参数3,参数4,参数5,参数6)$<br>
      </p></td>
    <td width="75%" class="tdbg">此标记调用某用户日志标题列表等信息。参数说明如下：<br>
      	　　参数1：userid。<br>
		　　参数2：调用日志条数。<br>
		　　参数3：日志标题长度，以字符为单位，超过部分显示“...”。<br>
		　　参数4：排序方法，为1按发表时间，为2按点击数，为3按回复数。<br>
		　　参数5：用户专题id，为0则调用该用户所有日志。<br>
		　　参数6：显示信息，1为显示发表日期，0为不显示。</td>
  </tr>

  此标记调用日志标题列表等信息。参数说明如下：
  <tr>
    <td height=23 class="tdbg">$show_class(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示系统分类列表。<br>
	　　参数1：横向显示时的每行条数，为0则竖向显示。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_bloger(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示博客列表。<br>
	　　参数1：横向显示时的每行条数，为0则竖向显示。
	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_treeclass(参数1)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示树型导航菜单，多级分类可展开关闭。<br>
	　　参数1：为log显示博客日志分类；<br>
	　　　　　 为user显示博客用户分类；<br>
	　　　　　 为photo显示博客相片类别；<br>
	　　　　　 为group显示博客群组分类。</font></td>
  </tr>
     <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_template(参数1)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示最新系统模板。<br>
	　　参数1：显示的数目</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_placard$</td>
    <td height=23 class="tdbg">此标记显示系统公告。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_count$</td>
    <td height=23 class="tdbg"> 此标记站点统计信息。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogupdate(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示日志更新排行列表。<br>
	　　参数1：调用条数。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_userlogin$</td>
    <td height=23 class="tdbg">此标记显示登录窗口。</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_userlogin_l$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示横向显示登录窗口。</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_comment(参数1,参数2)$</td>
    <td height=23 class="tdbg"> 此标记显示最新回复列表。<br>
　　参数1：调用条数；<br>
　　参数2：回复标题长度。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_subject(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示用户专题排行列表。<br>
	　　参数1：调用条数。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_bestblog(参数1)$</td>
    <td height=23 class="tdbg">此标记显示推荐博客。<br>
	　　参数1：调用条数。</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_hotblog(参数1,参数2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示热门博客,按照浏览量,回复数,留言数,被订阅数综合评定。<br>
    　　参数1：调用条数；<br>
	　　参数2：为0不显示博客头像，为1则显示博客头像。</font>
    	</td>
  </tr>
    <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_album(参数1,参数2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示用户的相册<br>
    　　参数1：调用条数；<br>
	　　参数2：为0以时间排序，为1则以访问量排序。</font>
    	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_pic(参数1,参数2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示用户的相片。<br>
    　　参数1：调用条数；<br>
	　　参数2：为0以时间排序，为1则以访问量排序，为2则以评论数排序。</font>
    	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_diggs(参数1,参数2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示DIGG。<br>
    　　参数1：调用条数；<br>
	　　参数2：为0以时间排序，为1则以被DIGG（推荐）数目多少排序。</font>
    	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_userdiggs(参数1,参数2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示用户被DIGG（推荐）的次数。<br>
    　　参数1：调用条数；<br>
	　　参数2：为0以被DIGG（推荐）的数目排序，为1则以用户注册时间排序。</font>
    	</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogstar$</td>
    <td height=23 class="tdbg">此标记调用最新博客之星。</td>
  </tr>
    <tr>
    <td height=23 class="tdbg">$show_blogstar2(参数1,参数2,参数3,参数4)$</td>
    <td height=23 class="tdbg">此标记调用最新博客之星。<br>
	　　参数1：调用数目；<br>
	　　参数2：每行显示数目；<br>
	　　参数3：图片宽度；<br>
	　　参数4：图片高度。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newblogger(参数1)$</td>
    <td height=23 class="tdbg">此标记显示最新注册用户。<br>
	　　参数1：调用条数。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_search(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示搜索表单。<br>
	　　参数1：为0横向显示，为1则竖向显示。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_cityblogger(参数1)$</td>
    <td height=23 class="tdbg">此标记显示同城博客搜索表单。<br>
	　　参数1：为0横向显示，为1则竖向显示。<font color="#FF0000">此标签不能用于副模板。</font>
    </td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newphoto(参数1,参数2,参数3,参数4)$</td>
    <td height=23 class="tdbg">此标签调用相册图片。<br>
	　　参数1：调用条数；<br>
	　　参数2：每行显示图片的数目；<br>
	　　参数3：图片宽度（单位：象素）；<br>
	　　参数4：图片高度（单位：象素）。</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_teams(参数1,参数2,参数3,参数4,参数5,参数6)$</font></td>
    <td height=23 class="tdbg"><font color=red> 此标记显示群组标志。<br/>
    　　参数1:调用类型 1- 最新创建/2-最活跃群组(贴数最多)/3-规模大(人数最多)；<br/>
    　　参数2:调用数目；<br/>
    　　参数3:题目显示长度；<br/>
    　　参数4:是否显示图标；<br/>
    　　参数5:图标宽度，不写则默认50象素；<br/>
	　　参数6:图标高度，不写则默认50象素。<br/></font>
    	</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_posts(参数1,参数2,参数3,参数4,参数5)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示群组新贴标志。<br/>
    　　参数1:显示某些群组的日志: 0 所有群组;如果是选择多个群组,则把群组ID用|分隔开,如1|2|8；<br/>
	　　参数2: 帖子数目；<br/>
	　　参数3:帖子主题显示字数；<br/>
	　　参数4:是否显示用户名 0/1；<br/>
	　　参数5:是否显示发帖时间 0/1。<br/></font>
    	</td>
  </tr>
  <tr align="left" >
    <td height=23 class="tdbg"><font color=red>$show_hottag(参数1,参数2,参数3,参数4)$</font></td>
    <td height=23 class="tdbg"> <font color=red>此标记显示TAG标志。<br/>
    　　参数1: 表现形式 1-列表形式,2-云图形式；<br/>
	　　参数2: 标签数目；<br/>
	　　参数3:已取消；<br/>
	　　参数4:每行显示数目。</font>
    	</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_friends$</td>
    <td height=23 class="tdbg"> 此标记显示友情链接。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> 此标记显示rss连接标志。</td>
  </tr>
  <tr>
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>副模板（对除index.asp外的其他系统页面有效，如list.asp,listblogger.asp文件等）：</font></strong></td>
  </tr>
  <tr>
    <td height=23 colspan="2" class="tdbg">包含所有主模板的标记<font color=#ff0000>($show_cityblogger(参数1)$ 标签除外)</font>，参数相同。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_list$</td>
    <td height=23 class="tdbg"> 重要，此标记显示其他系统次页面内容主体，不能去除，且只能在副模板使用。</td>
  </tr>
</table>
<br>
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 colspan="2" class="topbg"><strong>用户模板标记说明</strong>
  <tr>
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>主模板：</font><br>
      </strong>主模板为页面的主体部分，包括css样式设置等，建议在Dreamweave或Frontpage中编辑，完成后将代码copy进来。

    </td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_log$</td>
    <td height=23 class="tdbg"><strong>重要，此标记显示日志主体部分，包括评论等信息。</strong></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_subject$</td>
    <td height=23 class="tdbg"> 此标记以列表形式显示专题分类。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_subject_l$</td>
    <td height=23 class="tdbg"> 此标记横向显示专题分类。 </td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_login$</td>
    <td height=23 class="tdbg"> 此标记显示登录窗口。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_calendar$</td>
    <td height=23 class="tdbg"> 此标记显示日历。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_placard$ </td>
    <td height=23 class="tdbg">此标记显示用户公告。</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_photo$</td>
    <td height=23 class="tdbg"><font color=red>此标记显示博客相册最新相片，FLASH滚动播放。</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> 此标记显示最新日志列表。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> 此标记显示最新日志列表。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_comment$</td>
    <td height=23 class="tdbg"> 此标记显示最新回复列表。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newmessage$ </td>
    <td height=23 class="tdbg">此标记显示最新留言列表。</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_myfriend$</td>
    <td height=23 class="tdbg"><font color=red>此标记显示我的好友名单。</font></td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_mygroups$</td>
    <td height=23 class="tdbg"><font color=red>此标记显示加入的群组。</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_info$</td>
    <td height=23 class="tdbg"> 此标记显示Blog名称，统计信息等。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_search$</td>
    <td height=23 class="tdbg"> 此标记显示搜索表单。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_links$</td>
    <td height=23 class="tdbg"> 此标记显示链接信息。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogname$</td>
    <td height=23 class="tdbg">此标记显示用户blog名称，若名称为空则显示用户id。</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_blogurl$</font></td>
    <td height=23 class="tdbg"><font color=red>此标记显示博客的完整连接。</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> 此标记显示rss连接标志。</td>
  </tr>
  <tr>
<td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>副模板：</font><br>
      </strong>副模板为显示日志内容部分。包括日志标题，发表时间，日志内容等信息的版面设置。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_topic$</td>
    <td height=23 class="tdbg"> 此标记显示表情图标，专题名，是否加密，是否置顶，日志标题。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_emot$</td>
    <td height=23 class="tdbg">此标记仅显示标题表情图标。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_topictxt$</td>
    <td height=23 class="tdbg">此标记仅显示日志标题。</td>
  </tr>
  <tr>
    <td width="18%" height=23 class="tdbg">$show_loginfo$</td>
    <td width="82%" class="tdbg">此标记显示日志作者(评论者，留言者)，发表时间。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_author$</td>
    <td height=23 class="tdbg">此标记仅显示作者名(评论者，留言者)。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_addtime$</td>
    <td height=23 class="tdbg">此标记仅显示发表时间。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_logtext$</td>
    <td height=23 class="tdbg"> 此标记显示日志正文。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_more$</td>
    <td height=23 class="tdbg"> 此标记显示阅读全文(次数)，回复(次数)，引用链接。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogzhai$</td>
    <td height=23 class="tdbg">此标记显示加入到文摘的连接。</td>
  </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>