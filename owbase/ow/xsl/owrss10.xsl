<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"
                xmlns:rss="http://purl.org/rss/1.0/"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:h="http://www.w3.org/1999/xhtml"
                xmlns:hr="http://www.w3.org/2000/08/w3c-synd/#"
                xmlns:wiki="http://purl.org/rss/1.0/modules/wiki/"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                extension-element-prefixes="rdf rss dc h hr wiki msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:include href="owinc.xsl"/>

<xsl:template match="/">
    <xsl:apply-templates select="rdf:RDF"/>
</xsl:template>

<xsl:template match="rdf:RDF">
    <xsl:choose>
        <xsl:when test="rss:image">
            <xsl:apply-templates select="rss:image"/>
        </xsl:when>
        <xsl:otherwise>
            <xsl:apply-templates select="rss:channel"/>
        </xsl:otherwise>
    </xsl:choose>
    <xsl:apply-templates select="rss:item"/>
    <xsl:apply-templates select="rss:textinput"/>
</xsl:template>

<xsl:template match="rss:channel">
    <a href="{rss:link}" class="rss"><b><xsl:value-of select="rss:title"/></b></a><br />
</xsl:template>

<xsl:template match="rss:image">
    <a href="{rss:link}" class="rss"><img src="{rss:url}" alt="{rss:title}" border="0"/></a><br />
</xsl:template>

<xsl:template match="rss:item">
    <small>
    <xsl:if test="dc:date">
        <xsl:value-of select="ow:formatShortDateTime(string(dc:date))"/>
    </xsl:if>
    -
    <a href="{rss:link}" class="rss" title="{rss:description}"><xsl:value-of select="rss:title"/></a>
    <xsl:if test="wiki:status='new'">
        &#160;<span class="new">new</span>
    </xsl:if>
    <xsl:if test="wiki:status='deleted'">
        &#160;<span class="deprecated">deprecated</span>
    </xsl:if>
    <xsl:if test="wiki:diff">
        [<a href="{wiki:diff}" class="rss">diff</a>]
    </xsl:if>
    <xsl:if test="wiki:history">
        [<a href="{wiki:history}" class="rss">changes</a>]
    </xsl:if>
    &#160;
    <xsl:value-of select="dc:creator"/>
    <xsl:choose>
        <xsl:when test="dc:contributor/rdf:Description/@link">
            <a href="{dc:contributor/rdf:Description/@link}" class="rss" title="{dc:contributor/rdf:Description/@wiki:host}"><xsl:value-of select="dc:contributor/rdf:Description/rdf:value"/></a>
        </xsl:when>
        <xsl:otherwise>
            <xsl:choose>
                <xsl:when test="dc:contributor/rdf:Description/rdf:value/text()">
                    <xsl:value-of select="dc:contributor"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="dc:contributor/rdf:Description/@wiki:host"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:otherwise>
    </xsl:choose>
    </small>
    <br />
</xsl:template>

<xsl:template match="rss:textinput">
    <form action="{rss:link}" method="post">
        <xsl:value-of select="rss:title"/>:<br /><input type="text" name="{rss:name}"/>
    </form>
</xsl:template>

</xsl:stylesheet>