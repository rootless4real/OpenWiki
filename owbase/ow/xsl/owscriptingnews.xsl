<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:template match="/">
    <xsl:apply-templates select="scriptingNews"/>
</xsl:template>

<xsl:template match="scriptingNews">
    <xsl:apply-templates select="header"/>
    <xsl:apply-templates select="item"/>
    <xsl:if test="header/copyright">
        <xsl:value-of select="header/copyright"/>
        <br />
    </xsl:if>
</xsl:template>

<xsl:template match="header">
    <xsl:choose>
        <xsl:when test="string-length(imageUrl)&gt;0">
            <a href="{channelLink}" target="_blank"><img src="{imageUrl}" alt="{imageTitle}" border="0"/></a>
        </xsl:when>
        <xsl:otherwise>
            <a href="{channelLink}" target="_blank"><b><xsl:value-of select="channelTitle"/></b></a>
        </xsl:otherwise>
    </xsl:choose>
    <br />
</xsl:template>

<xsl:template match="item">
    <xsl:choose>
        <xsl:when test="link">
            <xsl:for-each select="link">

                <xsl:if test="position()=1">
                    <xsl:text disable-output-escaping="yes">&lt;ow:html&gt;&lt;![CDATA[</xsl:text>
                    <xsl:value-of select="substring-before(../text, linetext)" disable-output-escaping="yes" />
                    <xsl:text disable-output-escaping="yes">]]&gt;&lt;/ow:html&gt;</xsl:text>
                </xsl:if>

                <xsl:if test="position()&gt;1">
                    <xsl:text disable-output-escaping="yes">&lt;ow:html&gt;&lt;![CDATA[</xsl:text>
                    <xsl:value-of select="substring-after(substring-before(../text, linetext), ./preceding-sibling::*[position()=1]/linetext)" disable-output-escaping="yes" />
                    <xsl:text disable-output-escaping="yes">]]&gt;&lt;/ow:html&gt;</xsl:text>
                </xsl:if>

                <xsl:text disable-output-escaping="yes">&lt;ow:html&gt;&lt;![CDATA[</xsl:text>
                <a href="{url}" target="_blank"><xsl:value-of select="linetext" disable-output-escaping="yes"/></a>
                <xsl:text disable-output-escaping="yes">]]&gt;&lt;/ow:html&gt;</xsl:text>

                <xsl:if test="position()=last()">
                    <xsl:text disable-output-escaping="yes">&lt;ow:html&gt;&lt;![CDATA[</xsl:text>
                    <xsl:value-of select="substring-after(../text, linetext)" disable-output-escaping="yes" />
                    <xsl:text disable-output-escaping="yes">]]&gt;&lt;/ow:html&gt;</xsl:text>
                </xsl:if>
            </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
            <xsl:text disable-output-escaping="yes">&lt;ow:html&gt;&lt;![CDATA[</xsl:text>
            <xsl:value-of select="text" disable-output-escaping="yes" />
            <xsl:text disable-output-escaping="yes">]]&gt;&lt;/ow:html&gt;</xsl:text>
        </xsl:otherwise>
    </xsl:choose>
    <br />
    <br />
</xsl:template>

</xsl:stylesheet>