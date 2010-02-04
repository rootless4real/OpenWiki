<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"
                xmlns:ns="http://purl.org/rss/1.0/"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:ag="http://purl.org/rss/1.0/modules/aggregation/"
                xmlns:h="http://www.w3.org/1999/xhtml"
                xmlns:hr="http://www.w3.org/2000/08/w3c-synd/#"
                xmlns:wiki="http://purl.org/rss/1.0/modules/wiki/"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                extension-element-prefixes="rdf ns dc ag h hr wiki msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:include href="owinc.xsl"/>

<xsl:key name="items-by-timestamp" match="ns:item" use='substring-before(ag:timestamp, "T")' />

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

    <xsl:text disable-output-escaping="yes">&lt;ow:html&gt;&lt;![CDATA[</xsl:text>
        <table cellspacing="0" cellpadding="0" border="0" width="100%">
            <xsl:for-each select='ns:item[count(. | key("items-by-timestamp", substring-before(ag:timestamp, "T"))[1]) = 1]'>
                <xsl:sort select="ag:timestamp" order="descending" />

                <tr>
                    <td colspan="5"><br /><b><xsl:value-of select="ow:formatLongDate(string(ag:timestamp))"/></b></td>
                </tr>

                <xsl:for-each select='key("items-by-timestamp", substring-before(ag:timestamp, "T"))'>
                    <xsl:sort select="ag:timestamp" order="descending" />
<!--
                    <xsl:if test="position() mod 5 = 0">
                      <tr>
                        <td colspan="5">&#160;</td>
                      </tr>
                    </xsl:if>
-->
                    <tr>
                      <td width="1%" nowrap="nowrap">
                        <small>
                            <xsl:value-of select="ow:formatTime(string(ag:timestamp))"/>
                            &#160;
                            <xsl:if test="dc:date">
                                (<xsl:value-of select="ow:formatTime(string(dc:date))"/>)
                                &#160;
                            </xsl:if>
                        </small>
                      </td>
                      <td colspan="2" width="20%">
                        <xsl:choose>
                            <xsl:when test="ag:source/rdf:Description/@wiki:interwiki">
                                <a href="{ag:sourceURL}" class="rss" title="{ag:source}"><xsl:value-of select="ag:source/rdf:Description/@wiki:interwiki"/></a>:<a href="{ns:link}" class="rss" title="Version: {wiki:version}"><xsl:value-of select="ns:title"/></a>
                            </xsl:when>
                            <xsl:otherwise>
                                <a href="{ns:link}" class="rss" title="Version: {wiki:version}"><xsl:value-of select="ns:title"/></a>
                            </xsl:otherwise>
                        </xsl:choose>
                        <xsl:if test="wiki:status='new'">
                            &#160;<span class="new">new</span>
                        </xsl:if>
                        <xsl:if test="wiki:status='deleted'">
                            &#160;<span class="deprecated">deprecated</span>
                        </xsl:if>
                      </td>
                      <td nowrap="nowrap" width="5%">
                        &#160;
                        <xsl:if test="wiki:diff">
                            [<a href="{wiki:diff}" class="rss">diff</a>]&#160;
                        </xsl:if>
                        <xsl:if test="wiki:history">
                            [<a href="{wiki:history}" class="rss">changes</a>]&#160;
                        </xsl:if>
                      </td>
                      <xsl:choose>
                        <xsl:when test="ag:source/rdf:Description/@wiki:interwiki">
                          <td nowrap="nowrap" align="left">
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
                          </td>
                        </xsl:when>
                        <xsl:otherwise>
                          <td align="right">
                            <small><xsl:value-of select="dc:creator"/></small>
                            <a href="{ag:sourceURL}" class="rss"><xsl:value-of select="ag:source"/></a>
                          </td>
                        </xsl:otherwise>
                      </xsl:choose>
                    </tr>

                    <xsl:if test="string-length(ns:description) &gt; 0 and not(contains(ns:description, '&lt;a href='))">
                      <tr>
                        <td width="1%">&#160;</td>
                        <td width="1%">&#160;&#160;</td>
                        <td colspan="3" align="left">
                            <span class="comment"><xsl:value-of select="ns:description"/></span>
                        </td>
                      </tr>
                    </xsl:if>

                </xsl:for-each>
            </xsl:for-each>

        </table>
    <xsl:text disable-output-escaping="yes">]]&gt;&lt;/ow:html&gt;</xsl:text>

    <xsl:apply-templates select="ns:textinput"/>
</xsl:template>

<xsl:template match="ns:channel">
    <a href="{ns:link}" class="rss"><b><xsl:value-of select="ns:title"/></b></a><br />
</xsl:template>

<xsl:template match="ns:image">
    <a href="{ns:link}" class="rss"><img src="{ns:url}" alt="{ns:title}" border="0"/></a>
</xsl:template>

<xsl:template match="ns:textinput">
    <form action="{ns:link}" method="post">
        <xsl:value-of select="ns:title"/>:<br /><input type="text" name="{ns:name}"/>
    </form>
</xsl:template>

</xsl:stylesheet>