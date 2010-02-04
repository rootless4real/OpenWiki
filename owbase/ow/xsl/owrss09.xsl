<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"
                xmlns:ns="http://my.netscape.com/rdf/simple/0.9/"
                extension-element-prefixes="rdf ns"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:template match="/">
    <xsl:apply-templates select="rdf:RDF"/>
</xsl:template>

<xsl:template match="rdf:RDF">
    <xsl:choose>
        <xsl:when test="ns:image">
            <xsl:apply-templates select="ns:image"/>
        </xsl:when>
        <xsl:otherwise>
            <xsl:apply-templates select="ns:channel"/>
        </xsl:otherwise>
    </xsl:choose>
    <xsl:apply-templates select="ns:item"/>
    <xsl:apply-templates select="ns:textinput"/>
</xsl:template>

<xsl:template match="ns:channel">
    <a href="{ns:link}" title="{ns:description}" class="rss"><b><xsl:value-of select="ns:title"/></b></a><br />
</xsl:template>

<xsl:template match="ns:image">
    <a href="{ns:link}"><img src="{ns:url}" alt="{../ns:channel/ns:description}" border="0"/></a><br />
</xsl:template>

<xsl:template match="ns:item">
    <a href="{ns:link}" class="rss"><xsl:value-of select="ns:title"/></a><br />
</xsl:template>

<xsl:template match="ns:textinput">
    <form action="{ns:link}" method="post">
        <xsl:value-of select="ns:title"/>:<br /><input type="text" name="{ns:name}"/>
    </form>
</xsl:template>

</xsl:stylesheet>