<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding="utf-8" doctype-public="-//W3C//DTD XHTML 1.0 Transitional//EN"  doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" indent="yes" />
	<xsl:variable name="title" select="/rss/channel/title"/>
	<xsl:variable name="srclink" select="/rss/channel/link"/>
	<xsl:template match="/">
		<xsl:element name="html">
			<head>
				<title>
					<xsl:value-of select="$title"/>
				</title>
				<link rel="stylesheet" href="../../oblogstyle/rss.css" type="text/css" />
				<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
			</head>
			<xsl:apply-templates select="rss/channel"/>
		</xsl:element>
	</xsl:template>

	<xsl:template match="channel">
		<body>
			<div id="main">
				<!-- header -->
				<div id="header">
				<h1 title="logo">订阅<xsl:value-of select="$title"/></h1>
				<a href="{link}"><xsl:value-of select="$title"/></a>
				</div>
				<!-- header end -->
				<!-- contents -->
				<div id="contents">
					<!-- sidebar -->
					<div id="sidebar">
						<h4>订阅本站：</h4>
						<ul>
							<li><a href="http://www.oblog.cn/rss/?rss=http://{link}rss2.xml"><img src="http://www.oblog.cn/xml.jpg" border="0"/></a></li>
							<li><a href="{link}"><img src="../../images/xml.gif" border="0"/></a></li>
						</ul>
					</div>
					<!-- sidebar end -->
					<!-- content -->
					<div id="content">
						<xsl:apply-templates select="item"/>
					</div>
					<!-- content end -->
				</div>
				<!-- contents end -->
			</div>
			<!-- footer -->
			<div id="footer"></div>
			<!-- footer end -->
		</body>
	</xsl:template>

	<xsl:template match="item">
		<div class="log_title"><a href="{link}" title="{link}"><xsl:value-of select="title"/></a></div>
		<div class="log_time"><xsl:value-of select="author"/> 发表于 - <span><xsl:value-of select="pubDate"/></span></div>
		<div class="log_text"><table class="log_text_table"><tr><td><xsl:call-template name="outputContent"/></td></tr></table></div>
	</xsl:template>

	<xsl:template name="outputContent">
		<xsl:choose>
			<xsl:when test="xhtml:body" xmlns:xhtml="http://www.w3.org/1999/xhtml">
				<xsl:copy-of select="xhtml:body/*"/>
			</xsl:when>
			<xsl:when test="xhtml:div" xmlns:xhtml="http://www.w3.org/1999/xhtml">
				<xsl:copy-of select="xhtml:div"/>
			</xsl:when>
			<xsl:when test="content:encoded" xmlns:content="http://purl.org/rss/1.0/modules/content/">
				<xsl:value-of select="content:encoded" disable-output-escaping="yes"/>
			</xsl:when>
			<xsl:when test="description">
				<xsl:value-of select="description" disable-output-escaping="yes"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>
