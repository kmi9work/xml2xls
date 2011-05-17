<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
                xmlns:{1}="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItem{0}"
>
  <xsl:output method="xml" indent="yes"/>
  <xsl:template name="markchanges">
    <xsl:param name="cacheItem"/>
    <xsl:param name="currentVersionItem"/>
    <xsl:for-each select="$cacheItem">
      <!--<xsl:if test="name(.)='{1}:BaseItemVersion'">
          <{1}:BaseItem>
        </xsl:if>
        <xsl:if test="name(.)='{1}:AssortmentVersion'">
          <{1}:Assortment>
        </xsl:if>
        <xsl:if test="name(.)='{1}:PackagingItemVersion'">
          <{1}:PackagingItem>
        </xsl:if>-->
      <xsl:variable name="curCacheItem" select="."/>
      <xsl:variable name="gtin" select="./{1}:GTIN"/>
      <xsl:variable name="BI2" select="$currentVersionItem[{1}:GTIN=$gtin]"/>
      <xsl:if test="$BI2">
          <xsl:variable name="action">
            <xsl:value-of select="$BI2/{1}:ActionRequest"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="$action='DEL'">
              <xsl:copy>
                <xsl:attribute name="status">deleted</xsl:attribute>
                <!--<xsl:apply-templates select="$BI2/text()"/>-->
                <xsl:for-each select="$BI2/*">
                  <xsl:copy-of select="."/>
                </xsl:for-each>
              </xsl:copy>
            </xsl:when>
            <xsl:otherwise>
              <xsl:copy>
              <xsl:variable name="BI1" select="."/>
              <xsl:for-each select="./*">
                <xsl:variable name="node1" select="."/>
                <xsl:variable name="node2" select="$BI2/node()[name() = name($node1)]"/>
                <xsl:choose>
                  <xsl:when test="$node2">
                    <xsl:choose>
                      <xsl:when test="normalize-space($node1) = normalize-space($node2)">
                        <xsl:copy-of select="$node2"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:copy>
                          <xsl:apply-templates select="$node2/@*"/>
                          <xsl:attribute name="status">changed</xsl:attribute>
                          <xsl:apply-templates select="$node2/text()"/>
                          <xsl:for-each select="$node2/*">
                            <xsl:copy-of select="."/>
                          </xsl:for-each>
                        </xsl:copy>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:if test="$node1">
                      <xsl:copy>
                        <xsl:apply-templates select="$node1/@*"/>
                        <xsl:attribute name="status">deleted</xsl:attribute>
                        <xsl:apply-templates select="$node1/text()"/>
                        <xsl:for-each select="$node1/*">
                          <xsl:copy-of select="."/>
                        </xsl:for-each>
                      </xsl:copy>
                    </xsl:if>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
              <xsl:for-each select="$BI2/*">
                <xsl:variable name="node1" select="."/>
                <xsl:variable name="node2" select="$BI1/node()[name() = name($node1)]"/>
                <xsl:if test="not($node2)">
                  <xsl:copy>
                    <xsl:apply-templates select="$node1/@*"/>
                    <xsl:attribute name="status">added</xsl:attribute>
                    <xsl:apply-templates select="$node1/text()"/>
                    <xsl:for-each select="$node1/*">
                      <xsl:copy-of select="."/>
                    </xsl:for-each>
                  </xsl:copy>
                </xsl:if>
              </xsl:for-each>
              </xsl:copy>
            </xsl:otherwise>
          </xsl:choose>
      </xsl:if>
      <xsl:if test="not($BI2)">
        <xsl:variable name="curaction">
          <xsl:value-of select="$curCacheItem/{1}:ActionRequest"/>
        </xsl:variable>
        <xsl:if test="$curaction!='DEL'">
          <xsl:copy-of select="."/>
        </xsl:if>
      </xsl:if>
      <!--<xsl:if test="name(.)='{1}:BaseItemVersion'">
          </{1}:BaseItem>
        </xsl:if>
        <xsl:if test="name(.)='{1}:AssortmentVersion'">
          </{1}:Assortment>
        </xsl:if>
        <xsl:if test="name(.)='{1}:PackagingItemVersion'">
          </{1}:PackagingItem>
        </xsl:if>-->
    </xsl:for-each>
    <xsl:for-each select="$currentVersionItem">
      <xsl:variable name="gtin" select="./{1}:GTIN"/>
      <xsl:variable name="PI" select="$cacheItem[{1}:GTIN=$gtin]"/>
      <xsl:if test="name(.)='{1}:PackagingItemVersion'">
        <xsl:variable name="action">
          <xsl:value-of select="./{1}:ActionRequest"/>
        </xsl:variable>
        <xsl:if test="$action='ADD' and not($PI)">
          <xsl:copy>
            <xsl:apply-templates select="$PI/@*"/>
            <xsl:attribute name="status">added</xsl:attribute>
            <xsl:for-each select="./*">
              <xsl:copy-of select="."/>
            </xsl:for-each>
          </xsl:copy>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template match="/">
    <{1}:Item xmlns:{1}="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItem{0}">
      <xsl:variable name="cachedoc" select="document('{2}')"/>
      <xsl:variable name="versiondoc" select="document('{3}')"/>
      <xsl:if test="$cachedoc/{1}:Item/{1}:BaseItem/{1}:BaseItemVersion">
        <{1}:BaseItem>
          <xsl:call-template name="markchanges">
            <xsl:with-param name="cacheItem" select="$cachedoc/{1}:Item/{1}:BaseItem/{1}:BaseItemVersion"/>
            <xsl:with-param name="currentVersionItem" select="$versiondoc/{1}:Item/{1}:BaseItem/{1}:BaseItemVersion"/>
          </xsl:call-template>
        </{1}:BaseItem>
      </xsl:if>
      <xsl:if test="$cachedoc/{1}:Item/{1}:Assortment/{1}:AssortmentVersion">
        <{1}:Assortment>
          <xsl:call-template name="markchanges">
            <xsl:with-param name="cacheItem" select="$cachedoc/{1}:Item/{1}:Assortment/{1}:AssortmentVersion"/>
            <xsl:with-param name="currentVersionItem" select="$versiondoc/{1}:Item/{1}:Assortment/{1}:AssortmentVersion"/>
          </xsl:call-template>
        </{1}:Assortment>
      </xsl:if>
      <xsl:if test="$cachedoc/{1}:Item/{1}:PackagingItem/{1}:PackagingItemVersion">
        <{1}:PackagingItem>
          <xsl:call-template name="markchanges">
            <xsl:with-param name="cacheItem" select="$cachedoc/{1}:Item/{1}:PackagingItem/{1}:PackagingItemVersion"/>
            <xsl:with-param name="currentVersionItem" select="$versiondoc/{1}:Item/{1}:PackagingItem/{1}:PackagingItemVersion"/>
          </xsl:call-template>
        </{1}:PackagingItem>
      </xsl:if>
    </{1}:Item>
  </xsl:template>
</xsl:stylesheet>