<!--#include file="user_top.asp"-->
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="Help" class="FieldsetForm">
						<legend>用户后台帮助</legend>
						<h3><a name="1"></a>我的blog出现版面错乱，挪位怎么办？ </h3>
						<ul>
							<li>1、检查模板是否正常。</li>
							<li>2、检查日志是否有不标准的htm代码。</li>
							<li>3、对日志部分显示字数和排版进行微调，达到正常为止。</li>
							<li>4、建议使用部分显示标签进行首页部分显示日志的排版。</li>
						</ul>
						<h3><a name="2" id="2"></a>模板不小心改坏了怎么办？ </h3>
						<ul>
						重新选择一个默认模板即可。注：会将原来改过的模板覆盖掉，建议先备份模板。<br />
						</ul>
						<h3><a name="3" id="3"></a>日志发表了，但为什么首页没有显示？ </h3>
						<ul>
							<li>请用更新按钮，重新发布站点首页。</li>
						</ul>
						<h3><a name="4" id="4"></a>为什么无法上传文件？ </h3>
						<ul>
							<li>1、是您上传的文件大小超过了系统设定值，请压缩图片或文件。</li>
							<li>2、您的上传空间已满，请整理您的上传文件。</li>
							<li>3、您没有上传权限，请联系管理员。</li>
						</ul>
						<h3><a name="5" id="5"></a>一篇日志最多能写多少字？ </h3>
						<ul>
							<li>因受到数据库字段长度的限制，一篇文章请不要超过6万个英文字符，即：3万个中文字符。</li>
						</ul>
						<h3><a name="6" id="6"></a>可视化编辑器支持哪几种浏览器？ </h3>
						<ul>
							<li>oBlog集成的编辑器可以支持5.5以上版本的ie全系列浏览器、mozilla、firefox浏览器，在opear浏览器环境无法使用。</li>
						</ul>
						<h3><a name="7" id="7"></a>为什么我无法登录管理后台？ </h3>
						<ul>
							<li>1、请确认用户名和密码输入正确。</li>
							<li>2、登录系统需要cookies环境，请检查浏览器的cookies是否关闭。</li>
							<li>3、请联系系统管理员。</li>
						</ul>
						<h3><a name="9" id="9"></a>如何修改我博客的个性模板？ </h3>
						<ul>
							<li>请先选择一个喜欢的默认模板，然后选择<strong>修改模板</strong>菜单进行操作。</li>
							<li>建议将代码拷贝到Dreamweaver 或者Frontpage编辑</li>
						<li class="red">注意：本系统分为主模板和副模板，主模板为网站的整体结构，副模板的修改只对日志主体部分起作用，也就是对标签$show_log$起作用，具体调用标签如下，您可以灵活运用，做个个性的模板（建议修改前先备份模板）</li>
						<li>用户模板标记说明 </li>
						<li>
							<ul>
								<li><strong>主模板：</strong></li>
								<li>$show_log$ 重要，此标记显示日志主体部分，包括评论等信息。</li>
								<li>$show_placard$ 此标记显示用户公告。</li>
								<li>$show_calendar$ 此标记显示日历。</li>
								<li>$show_newblog$ 此标记显示最新日志列表。</li>
								<li>$show_comment$ 此标记显示最新回复列表。</li>
								<li>$show_subject$ 此标记显示专题分类。</li>
								<li>$show_subject_l$ 此标记横向显示专题分类。</li>
								<li>$show_newblog$ 此标记显示最新日志列表。</li>
								<li>$show_newmessage$ 此标记显示最新留言列表。</li>
								<li>$show_info$ 此标记显示Blog名称，统计信息等。</li>
								<li>$show_login$ 此标记显示登录窗口。</li>
								<li>$show_links$ 此标记显示链接信息。</li>
								<li>$show_blogname$ 此标记显示用户blog名称，若名称为空则显示用户id。</li>
								<li>$show_search$ 此标记显示搜索表单。</li>
								<li>$show_xml$ 此标记显示rss连接标志。</li>
								<li>$show_myfriend$ 此标记显示我的好友。</li>
								<li>$show_mygroup$ 此标记显示我的群组。</li>
								<li>$show_photo$ 此标记显示相册。</li>
								<li>$show_blogurl$ 此标记显示博客连接地址。</li>
							</ul>
							<ul>
								<li><strong>副模板：</strong></li>
								<li>$show_topic$ 此标记显示日志题目。</li>
								<li>$show_loginfo$ 此标记显示日志作者，发表时间等信息。</li>
								<li>$show_logtext$ 此标记显示日志正文。</li>
								<li>$show_more$ 此标记显示阅读全文，引用等链接。</li>
								<li>$show_emot$ 此标记仅显示显示表情图标。</li>
								<li>$show_author$ 此标记仅显示作者名。</li>
								<li>$show_addtime$ 此标记仅显示发表时间。</li>
								<li>$show_topictxt$ 此标记仅显示日志标题。</li>
							</ul>

						<li></li>
						</li>
						<li class="red">注意：若不小心将模板改坏，可以重新选择默认模板进行恢复。</li>
						</ul>
						<h3><a name="10" id="10"></a>如何修改我<%=oblog.CacheConfig(69)%>的个性模板？ </h3>
							<ul><li><ul>
<li><strong><%=oblog.CacheConfig(69)%>主模版标签:</strong></li>
<li>$group_id$ <%=oblog.CacheConfig(69)%>ID</li>
<li>$group_posts$ 最新文章</li>
<li> $group_ico$  <%=oblog.CacheConfig(69)%>标记图片 </li>
<li> $group_url$ <%=oblog.CacheConfig(69)%>访问地址 </li>
<li>$group_guide$ 导航链接</li>
<li> $group_name$ <%=oblog.CacheConfig(69)%>名字 </li>
<li> $group_creater$ <%=oblog.CacheConfig(69)%>创建人 </li>
<li> $group_bottom$ 版权标识</li>
<li> $group_comments$ 最近评论  </li>
<li>$group_placard$ 公告</li>
<li> $group_links$ 友情链接 </li>
<li> $group_info$ <%=oblog.CacheConfig(69)%>信息</li>
<li> $group_bestuser$ 活跃用户</li>
<li> $group_newuser$ 最新加入用户</li>
<li>  $group_admin$ 管理员信息</li>
<li> $group_bestposts$ 精华帖子 </li>
<li>$group_photo$ <%=oblog.CacheConfig(69)%>相片 </li></ul>
							</li>
							<li>
							<ul>
							
							<li><strong><%=oblog.CacheConfig(69)%>副模版标签</strong></li>
<li> $group_list$ 内容标签 </li>
<li> $group_post_title$ 帖子标题 </li>
<li>  $group_content$ 帖子内容</li>
<li>  $group_post_userico$ 作者头像</li>
<li>  $group_post_user$ 帖子作者 </li>
<li>  $group_post_time$ 发布时间 </li>
<li>  $group_post_content$ 帖子正文 </li>
<li>  $group_post_id$ 帖子ID </li>
<li> $group_post_replys$ 回复按钮 </li>
<li>  $group_post_userurl$ 帖子作者地址 </li>
<li>   $group_post_high$ 帖子楼层  </li>
<li> $group_post_m$帖子操作导航 </li>
<li></li>
							</ul>
							</li>
							<li></li>
							</ul>
						<h3><a name="tb">什么是trackback ping(引用通告)？</a></h3>
						<ul>
							<li><strong>一、“引用通告”是什么？</strong></li>
							<li>　　“引用通告”简单的说，就是如果你写的文章是根据其他人Blog中的文章而做出的延伸或评论，你可以通知对方你针对他的文章写了东西，这就需要用到引用通告。</li>
							<li>　　在以往我们的经验当中，您对他人日志文章的评论只能在他人文章后通过回复进行，这样做让我们只能在别人的地盘上活动，而不能自己掌握自己的发表的言论，这就带来一些麻烦。</li>
							<li>　　第一，您发表在他人文章后面的评论您没有办法再进行修改。如果张三的Blog上有一篇我感兴趣的文章，您在这篇文章下发表自己的评论，但您的评论只能存在于张三的Blog上，您无法再修改增删这篇评论。</li>
							<li>　　第二，您希望获得张三关注的文章只能采取在自己的Blog中写了一篇和张三类似的文章，您希望张三能来看一看我写的这篇文章，这时您就必须到张三的Blog的那篇文章下发一篇回复，同时把您想让他看的那篇文章地址贴上去。
							有了引用通告，您就可以完全不需要这样麻烦了，完全可以在自己的BLOG里进行操作，彻底享受自己掌握自己言论的主动权。</li>
							<li>　　通过引用通告，您就可以在自己的Blog中发表文章，同时把自己这篇文章的地址发到张三的Blog的那篇文章上去。
							在自己的地盘上引用张三的文章，然后通知他，“嘿，您的文章被我评论了”，这就是“引用通告”。</li>
							<li>　　同样的，当别人引用您的文章的时候，系统也可以接收对方的请求并进行记录，这样您可以查看来源地址，看看对方是如何评论您的文章的。</li>
							<li></li>
						</ul>
						<ul>
							<li><strong>二、如何使用“引用通告”</strong></li>
							<li>　　1、找到你要评论的Blog日志的“引用通告”地址，一般在日志下方有“引用通告”项，点击进入后可以看到地址，或者有的Blog直接在日志下方显示，把地址复制下来；</li>
							<li>　　2、进入自己的Blog发表新日志，在下方有“引用通告”栏目，将要评论的Blog的“引用通告”地址粘贴在这里。每行您可以粘帖一个。</li>
							<li>　　3、发表自己的日志后，系统会自动向目标地址发送引用申请，之后你会在要评论的Blog中看到你的Blog日志地址。</li>
							<li></li>
							<li class="blue">引用通告并不神秘和复杂，您可以先和朋友互相试验一下，相信您很快就会发现引用通告给您带来的全新感受。</li>
						</ul>
						<h3><a name="tag">什么是Tag？</a></h3>
						<ul>
							<li><strong>一、什么是标签（TAG）？</strong></li>
							<li>　　简单的说,标签就是一篇文章的"关键词"。您可以将日志文章或者照片，选择一个或多个词语（标签）来标记，这样一来，凡是我们博客网站上使用该词语的文章自动成为一个列表显示。</li>
							<li></li>
						</ul>
						<ul>
							<li><strong>二、使用标签的好处：</strong></li>
							<li>　　1、您添加标签的文章就会被直链接到网站相应标签的页面，这样浏览者在访问相关标签时，就有可能访问到您的文章，增加了您的文章被访问的机会。</li>
							<li>　　2、您可以很方便地查找到与您使用了同样标签的文章，延伸您文章的视野；可以方便地查找到与您使用了同样标签的作者，作为志同道合的朋友，您可以将他们加为好友或友情博客，扩大您的朋友圈。</li>
							<li>　　3、增加标签的方式完全由您自主决定，不受任何的限制，不用受网站系统分类和自己原有日志分类的限制，便于信息的整理、记忆和查找。</li>
							<li></li>
						</ul>
						<ul>
							<li><strong>三、如何使用标签?</strong></li>
							<li>　　例如：您写了一篇到北京旅游的文章，按照文章提到的内容，您可以给这篇文章加上：</li>
							<li>　　北京旅游,天安门,长城,故宫</li>
							<li>　　等几个标签，当浏览者想搜索关于长城的文章时，浏览者会点击标签：长城，从而看到所有关于长城的文章，方便了浏览者查找日志，同时您也可用此方法找到和您同样喜欢的人，以便一起相互交流等等。</li>
							<li></li>
						</ul>
						<ul>
							<li><strong>四．如何添加“好”的Tag？</strong></li>
							<li>　　1． Tag应该要能够体现出自己的特色，并且是大家经常采用的熟悉的词语。</li>
							<li>　　2．用词尽量简单精炼,词语字数不要太长，两三个字的词语就可以了，尽量是有意义的词汇，不要使用一些只作为装饰的符号，如｛｝等。</li>
							<li>　　3．不要使用一些语义比较弱的词汇，如“我的家”，“图片”等。</li>
						</ul>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>