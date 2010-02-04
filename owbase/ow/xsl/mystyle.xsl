<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:variable name="brandingText">OpenWiki, the post-it note of the web.</xsl:variable>

<xsl:template name="head">
  <head>
  <meta http-equiv="Content-Type" content="text/html; charset={@encoding};" />
  <meta name="keywords" content="wiki, openwiki, wikiwiki, wiki wiki, interwiki, post-it note, collaborative web, open source, ASP, IIS, XML, XHTML, RSS, dotnet, weblog, log, Lawrence Pit, LawrencePit"/>
  <meta name="description" content="OpenWiki - The post-it note of the web."/>
  <meta name="ROBOTS" content="INDEX,FOLLOW"/>
  <meta name="MSSmartTagsPreventParsing" content="true"/>
  <title><xsl:value-of select="ow:title"/> - <xsl:value-of select="ow:page/ow:link"/></title>
  <link rel="stylesheet" type="text/css" href="ow/css/ow.css" />
  </head>
</xsl:template>

<xsl:template name="brandingImage">
    <a href="{/ow:wiki/ow:frontpage/@href}"><img src="{/ow:wiki/ow:imagepath}/logo.gif" align="right" border="0" alt="OpenWiki" /></a>
</xsl:template>

<xsl:template name="poweredBy">
    <a href="http://openwiki.com"><img src="{/ow:wiki/ow:imagepath}/poweredby.gif" width="88" height="31" border="0" alt="" /></a>
</xsl:template>

<xsl:template name="validatorButtons">
    <a href="http://validator.w3.org/check/referer"><img src="{/ow:wiki/ow:imagepath}/valid-xhtml10.gif" alt="Valid XHTML 1.0!" width="88" height="31" border="0" /></a>
    <a href="http://jigsaw.w3.org/css-validator/validator?uri={/ow:wiki/ow:location}ow.css"><img src="{/ow:wiki/ow:imagepath}/valid-css.gif" alt="Valid CSS!" width="88" height="31" border="0" /></a>
</xsl:template>


</xsl:stylesheet>