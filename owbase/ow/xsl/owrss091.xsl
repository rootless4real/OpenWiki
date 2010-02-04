<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:template match="/">
    <xsl:apply-templates select="rss"/>
</xsl:template>

<xsl:template match="rss">
    <xsl:apply-templates select="channel"/>
</xsl:template>

<xsl:template match="channel">
    <xsl:choose>
        <xsl:when test="image">
            <xsl:apply-templates select="image"/>
        </xsl:when>
        <xsl:otherwise>
            <a href="{link}" title="{description}" class="rss"><b><xsl:value-of select="title"/></b></a><br />
        </xsl:otherwise>
    </xsl:choose>
    <xsl:apply-templates select="item"/>
</xsl:template>

<xsl:template match="image">
    <a href="{link}"><img src="{url}" alt="{description}" border="0"/></a><br />
</xsl:template>

<xsl:template match="item">
    <xsl:choose>
        <xsl:when test="title">
            <a href="{link}" title="{description}" class="rss"><xsl:value-of select="title"/></a>
        </xsl:when>
        <xsl:otherwise>
            <xsl:text disable-output-escaping="yes">&lt;ow:html&gt;&lt;![CDATA[</xsl:text>
                <p>
                    <xsl:value-of select="description" disable-output-escaping="yes"/>
                </p>
            <xsl:text disable-output-escaping="yes">]]&gt;&lt;/ow:html&gt;</xsl:text>
        </xsl:otherwise>
    </xsl:choose>
    <br />
</xsl:template>

</xsl:stylesheet>