addLoadEvent( function() {var manager = new dbxManager('postmeta');} );

addLoadEvent( function()
{
	//create new docking boxes group
	var meta = new dbxGroup(
		'content_li', 		// container ID [/-_a-zA-Z0-9/]
		'vertical', 	// orientation ['vertical'|'horizontal']
		'10', 			// drag threshold ['n' pixels]
		'no',			// restrict drag movement to container axis ['yes'|'no']
		'10', 			// animate re-ordering [frames per transition, or '0' for no effect]
		'yes', 			// include open/close toggle buttons ['yes'|'no']
		'closed', 		// default state ['open'|'closed']
		'打开', 		// word for "open", as in "open this box"
		'关闭', 		// word for "close", as in "close this box"
		'按下鼠标并拖动以移动此框', // sentence for "move this box" by mouse
		'点击以%固定%此框', // pattern-match sentence for "(open|close) this box" by mouse
		'使用箭头键移动此框', // sentence for "move this box" by keyboard
		'，或点击回车键%固定%它',  // pattern-match sentence-fragment for "(open|close) this box" by keyboard
		'%mytitle%  [%dbxtitle%]' // pattern-match syntax for title-attribute conflicts
		);

	// Boxes are closed by default. Open the Category box if the cookie isn't already set.
	var catdiv = document.getElementById('categorydiv');
	if ( catdiv ) {
		var button = catdiv.getElementsByTagName('A')[0];
		if ( dbx.cookiestate == null && /dbx\-toggle\-closed/.test(button.className) )
			meta.toggleBoxState(button, true);
	}

	var advanced = new dbxGroup(
		'advancedstuff', 		// container ID [/-_a-zA-Z0-9/]
		'vertical', 		// orientation ['vertical'|'horizontal']
		'10', 			// drag threshold ['n' pixels]
		'yes',			// restrict drag movement to container axis ['yes'|'no']
		'10', 			// animate re-ordering [frames per transition, or '0' for no effect]
		'yes', 			// include open/close toggle buttons ['yes'|'no']
		'closed', 		// default state ['open'|'closed']
		'打开', 		// word for "open", as in "open this box"
		'关闭', 		// word for "close", as in "close this box"
		'按下鼠标并拖动以移动此框', // sentence for "move this box" by mouse
		'点击以%固定%此框', // pattern-match sentence for "(open|close) this box" by mouse
		'使用箭头键移动此框', // sentence for "move this box" by keyboard
		'，或点击回车键%固定%它',  // pattern-match sentence-fragment for "(open|close) this box" by keyboard
		'%mytitle%  [%dbxtitle%]' // pattern-match syntax for title-attribute conflicts
		);
});
