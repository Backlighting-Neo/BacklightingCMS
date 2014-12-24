<?xml version="1.0" encoding="GB2312" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:dc="http://purl.org/dc/elements/1.1/" version="1.0">
<xsl:template match="/rss">
<html>
<head>
<style type="text/css" rel="stylesheet">
body { margin-top:10px; margin-bottom:25px; background:#d4d0c8;text-align:center; font-family:Verdana,Simsun; font-size: 12px; line-height: 1.45em; }
#feedMain { margin:0px auto; width:860px; text-align:left; word-break: break-all;word-wrap: break-word;overflow: hidden;}
p { padding-top: 0px; margin-top: 0px; }
td { font-family:Verdana,Simsun; font-size: 12px; line-height: 1.45em; word-break: break-all;word-wrap: break-word;overflow: hidden;}
h1 { font-size: 16px; padding-bottom: 0px; margin-bottom: 0px; }
h2 { font-size: 14px; margin-bottom: 0px; }

.gray { color: #808080; }
.t { font-size:14px; color: #000000; font-weight:bold; }
#feedBody2 {background:#fff;border:1px solid THreeDShadow;margin:0px auto;padding:1px 10px;}
#feedBody hr {border: 0px solid #3165c6;border-top-width: 1px;height: 0px;padding: 0px;margin:0px;}
#footer {margin:auto;text-align:center;color:#111;}
#footer hr   {border: 0px dashed #000;border-top-width: 1px;height: 0px;margin: 8px 0px 8px 0px;padding: 0px;width:80%;}

a { color:#ff3300; text-decoration: underline; }
a:hover { color:#ff0000; text-decoration: underline; }

h2 a { color:#3165c6; text-decoration: none; }
h2 a:hover { color:#3e80fa; text-decoration: none; }

#footer a { color:#111; text-decoration: none; }
#footer a:hover { color:#3e80fa; text-decoration: underline; }

#feedHeader {
  margin:0px auto;padding:1px 10px;
  background:#ffffc6;
  border:1px solid THreeDShadow;
}

#feedBody {
  border: 1px solid THreeDShadow;
  padding: 3em;
  -moz-padding-start: 30px;
  margin: 2em auto;
  background: #fff;
}

#feedTitleLink {
  float: right;
  -moz-margin-start: .6em;
  -moz-margin-end: 0;
  margin-top: 0;
  margin-bottom: 0;
}

a[href] img {
  border: none;
}

#feedTitleContainer {
  -moz-margin-start: 0;
  -moz-margin-end: .6em;
  margin-top: 0;
  margin-bottom: 0;
}

#feedTitleImage {
  -moz-margin-start: .6em;
  -moz-margin-end: 0;
  margin-top: 0;
  margin-bottom: 0;
  max-width: 300px;
  max-height: 150px;
}

div#feedTitleContainer h1 {
  font-size: 160%;
  border-bottom: 2px solid ThreeDLightShadow;
  margin: 0 0 .2em 0;
}

div#feedTitleContainer h2 {
  color: ThreeDDarkShadow;
  font-size: 110%;
  font-weight: normal;
  margin: 0 0 .6em 0;
}
</style>
<title><xsl:value-of select="channel/title" /></title>
</head>
<body>
<div id="feedMain">
	<div id="feedHeader">
		<h1><a href="{channel/titlelink}" target="_blank"><xsl:value-of select="channel/title" /></a></h1>
		<p class="gray">说明：您可以使用任何 RSS 阅读器或聚合器来订阅本站提供的“聚合新闻服务”（RSS），方便的预览本网站的更新。RSS 允许您订阅多个源，并自动将信息组合到一个列表中。</p>
		<p><a href="{channel/Currentlink}" target="_blank">点击订阅本 RSS 更新</a> <font class="gray">[您需要事先安装支持单击订阅功能的 RSS 聚合器]</font>
		<a href="#" onmouseover="window.status='有关 RSS 的详细知识';return true;" onmouseout="window.status='';return true;">有关 RSS 的详细知识</a></p>
	</div>
	<br />
	<div id="feedBody">
	     <div id="feedTitle">
			<a id="feedTitleLink" title="转到 {channel/description}" href="/">
				<img id="feedTitleImage" src="{channel/image/url}" border="0"/>
			</a>
		<div id="feedTitleContainer">
			<h1 id="feedTitleText"><xsl:value-of select="channel/title" /></h1>
			<h2 id="feedSubtitleText"><xsl:value-of select="channel/description" /></h2>
		</div>
	</div>
		<xsl:apply-templates select="channel/item" />
	</div>
	<div id="footer">
		<hr />(C)<font face="Verdana, Arial, Helvetica, sans-serif"><b>actcms<font color="#CC0000">.com</font></b></font></div>
</div>
</body>
</html><script src="/count.asp" type="text/javascript"></script>
</xsl:template>

<xsl:template match="item">
	<h2><a href="{link}" target="_blank"><xsl:value-of select="title" /></a></h2>
	<hr />
	<font class="gray">更新日期: <xsl:value-of select="pubDate" /></font>
	<p>
	<xsl:value-of select="content" />
	<xsl:value-of select="description" disable-output-escaping="yes" />
	<br /><a href="{link}" target="_blank">阅读全文</a>
	<br />
	<br />
	  <xsl:if test="category != ''">
		<font color="gray">类别：<xsl:value-of select="category" /></font>
		<br />
	  </xsl:if>
	</p>
</xsl:template>
</xsl:stylesheet>