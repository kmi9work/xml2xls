<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
                xmlns:sinfos="http://schemas.sinfos.de/TradeItemMessages/1.2.0/TradeItemMessage"
				xmlns:fnf_{1}="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_{2}"
                xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
>
  <xsl:output method="xml" version="1.0" indent="yes" omit-xml-declaration="no" standalone="yes"/>
  <xsl:template name="for.loop">
    <xsl:param name="num">1</xsl:param>
    <!-- param has initial value of1 -->
    <xsl:if test="not($num = 47)">
      <s:Column s:Width="100"/>
      <xsl:call-template name="for.loop">
        <xsl:with-param name="num">
          <xsl:value-of select="$num + 1"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>
  <xsl:template name="markcellchanges">
    <xsl:param name="node"/>
    <xsl:param name="item"/>
    <xsl:param name="styleID"/>
    <xsl:choose>
      <xsl:when test="$item/@status='changed' or $node/@status='changed'">
        <xsl:attribute name="s:StyleID">
          <xsl:value-of select="concat(substring($styleID,1,1), '1',substring($styleID,2))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$item/@status='added' or $node/@status='added'">
        <xsl:attribute name="s:StyleID">
          <xsl:value-of select="concat(substring($styleID,1,1), '2',substring($styleID,2))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$item/@status='deleted' or $node/@status='deleted'">
        <xsl:attribute name="s:StyleID">
          <xsl:value-of select="concat(substring($styleID,1,1),'3',substring($styleID,2))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="$styleID">
          <xsl:attribute name="s:StyleID">
            <xsl:value-of select="$styleID"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="for.item">
    <xsl:param name="item" />
    <xsl:param name="biGTIN"/>
    <xsl:param name="biPackagingFNF"/>
    <xsl:param name="BI"/>
    <s:Row>
      <xsl:variable name="status" select="$item/@status"/>
      <xsl:choose>
        <xsl:when test="$status='changed'">
          <xsl:attribute name="ss:StyleID">s56</xsl:attribute>
        </xsl:when>
        <xsl:when test="$status='deleted'">
          <xsl:attribute name="ss:StyleID">s58</xsl:attribute>
        </xsl:when>
        <xsl:when test="$status='added'">
          <xsl:attribute name="ss:StyleID">s57</xsl:attribute>
        </xsl:when>
      </xsl:choose>
      <s:Cell >
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:GTIN"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s38</xsl:text>
          </xsl:with-param>  
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:value-of select="$item/fnf_{1}:GTIN"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$biGTIN"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s38</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:value-of select="$biGTIN"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s39</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s39</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:InternalItemIDofSupplier"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s40</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:InternalItemIDofSupplier">
            <s:Data s:Type="Number">
              <xsl:value-of select="$item/fnf_{1}:InternalItemIDofSupplier"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s41</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}']"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s42</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}']">
              <xsl:value-of select="substring($item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}'], 0, 20)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring($BI/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}'], 0, 20)"/>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}']"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s42</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}']">
              <xsl:value-of select="substring($item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}'], 20, 20)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring($BI/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}'], 20, 20)"/>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s43</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s44</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:NumberOfBaseUnitContained"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s45</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:NumberOfBaseUnitContained">
            <s:Data s:Type="Number">
              <xsl:value-of select="$item/fnf_{1}:NumberOfBaseUnitContained"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:CountryOfOrigin"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s21</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:CountryOfOrigin">
              <mapTo list="countries" firstcell="code" secondcell="withcode">
                <xsl:value-of select="$item/fnf_{1}:CountryOfOrigin"/>
              </mapTo>
            </xsl:when>
            <xsl:otherwise>
              <mapTo list="countries" firstcell="code" secondcell="withcode">
                <xsl:value-of select="$BI/fnf_{1}:CountryOfOrigin"/>
              </mapTo>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s47</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:text>0001</xsl:text>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s47</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:text>0001</xsl:text>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:GrossWeightFNFTSS"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s48</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:GrossWeightFNFTSS">
            <s:Data s:Type="Number">
              <xsl:value-of select="$item/fnf_{1}:GrossWeightFNFTSS"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemLengthDepth"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s47</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:MeasurementTradeItemLengthDepth">
            <s:Data s:Type="Number">
              <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemLengthDepth"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemWidth"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s47</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:MeasurementTradeItemWidth">
            <s:Data s:Type="Number">
              <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemWidth"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemHeight"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s42</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:MeasurementTradeItemHeight">
            <s:Data s:Type="Number">
              <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemHeight"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s44</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s49</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s44</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:PLUDescriptionML/fnf_{1}:PLUDescriptionML[@language='{0}']"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s50</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:PLUDescriptionML/fnf_{1}:PLUDescriptionML[@language='{0}']">
              <xsl:value-of select="$item/fnf_{1}:PLUDescriptionML/fnf_{1}:PLUDescriptionML[@language='{0}']"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$BI/fnf_{1}:PLUDescriptionML/fnf_{1}:PLUDescriptionML[@language='{0}']"/>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s51</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s51</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <!--<s:Data s:Type="Number">
        </s:Data>-->
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s52</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <!--<s:Data s:Type="DateTime">
        </s:Data>-->
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s52</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <!--<s:Data s:Type="DateTime">
        </s:Data>-->
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s42</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:text>RUB</xsl:text>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s51</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s51</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s52</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s52</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s42</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:text>RUB</xsl:text>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s50</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s50</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s50</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell mapTo="packaging">
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:PackagingTypeFNF"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s21</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$BI/fnf_{1}:PackagingTypeFNF">
              <xsl:value-of select="$BI/fnf_{1}:PackagingTypeFNF"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>-</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell mapTo="packaging">
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:PackagingTypeFNF"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s21</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:PackagingTypeFNF">
              <xsl:value-of select="$item/fnf_{1}:PackagingTypeFNF"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>-</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:ContentUnitOfMeasure"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s21</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:ContentUnitOfMeasure">
              <mapTo list="measurement" firstcell="code" secondcell="globus">
                <xsl:value-of select="$item/fnf_{1}:ContentUnitOfMeasure"/>
              </mapTo>
            </xsl:when>
            <xsl:otherwise>
              <mapTo list="measurement" firstcell="code" secondcell="globus">
                <xsl:value-of select="$BI/fnf_{1}:ContentUnitOfMeasure"/>
              </mapTo>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:CertificatesOfQuality"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s42</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:CertificatesOfQuality/fnf_{1}:certificatesOfQuality/fnf_{1}:certificateOfQuality/fnf_{1}:certificateOfQualityNumber"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:CertificatesOfQuality"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s52</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:CertificatesOfQuality/fnf_{1}:certificatesOfQuality/fnf_{1}:certificateOfQuality/fnf_{1}:certificateOfQualityStartDate">
            <s:Data s:Type="DateTime">
              <xsl:value-of select="$item/fnf_{1}:CertificatesOfQuality/fnf_{1}:certificatesOfQuality/fnf_{1}:certificateOfQuality/fnf_{1}:certificateOfQualityStartDate"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:CertificatesOfQuality"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s52</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$item/fnf_{1}:CertificatesOfQuality/fnf_{1}:certificatesOfQuality/fnf_{1}:certificateOfQuality/fnf_{1}:certificateOfQualityEndDate">
            <s:Data s:Type="DateTime">
              <xsl:value-of select="$item/fnf_{1}:CertificatesOfQuality/fnf_{1}:certificatesOfQuality/fnf_{1}:certificateOfQuality/fnf_{1}:certificateOfQualityEndDate"/>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String">
            </s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s53</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$BI/fnf_{1}:MinimumDurabilityFromArrival">
            <s:Data s:Type="Number">
              <xsl:choose>
                <xsl:when test="$item/fnf_{1}:MinimumDurabilityFromArrival">
                  <xsl:value-of select="$item/fnf_{1}:MinimumDurabilityFromArrival"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
                </xsl:otherwise>
              </xsl:choose>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String"></s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:Content"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s54</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:choose>
          <xsl:when test="$BI/fnf_{1}:Content">
            <s:Data s:Type="Number">
              <xsl:choose>
                <xsl:when test="$item/fnf_{1}:Content">
                  <xsl:value-of select="$item/fnf_{1}:Content"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$BI/fnf_{1}:Content"/>
                </xsl:otherwise>
              </xsl:choose>
            </s:Data>
          </xsl:when>
          <xsl:otherwise>
            <s:Data s:Type="String"></s:Data>
          </xsl:otherwise>
        </xsl:choose>
      </s:Cell>
      <s:Cell mapTo="nds">
         <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:VAT_{0}"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s55</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:VAT_{0}">
              <xsl:value-of select="$item/fnf_{1}:VAT_{0}"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$BI/fnf_{1}:VAT_{0}"/>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
        </s:Data>
      </s:Cell>
    </s:Row>
  </xsl:template>
  <xsl:template match="/">
    <xsl:for-each select="fnf_{1}:Item">
      <xsl:variable name="biGTIN" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion/fnf_{1}:GTIN"/>
      <xsl:variable name="biGTINA" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion/fnf_{1}:GTIN"/>
      <xsl:variable name="BI" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion" />
      <xsl:variable name="BA" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion" />
      <xsl:if test="$biGTIN">
        <xsl:variable name="biPackagingFNF" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion/fnf_{1}:PackagingTypeFNF"/>
        <xsl:call-template name="for.item">
          <xsl:with-param name="item" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion"></xsl:with-param>
          <xsl:with-param name="biGTIN" select="$biGTIN"></xsl:with-param>
          <xsl:with-param name="biPackagingFNF" select="$biPackagingFNF"></xsl:with-param>
          <xsl:with-param name="BI" select="$BI"></xsl:with-param>
        </xsl:call-template>
        <xsl:for-each select="fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion">
          <xsl:call-template name="for.item">
            <xsl:with-param name="item" select="."></xsl:with-param>
            <xsl:with-param name="biGTIN" select="$biGTIN"></xsl:with-param>
            <xsl:with-param name="biPackagingFNF" select="$biPackagingFNF"></xsl:with-param>
            <xsl:with-param name="BI" select="$BI"></xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="$biGTINA">
        <xsl:variable name="biPackagingFNFA" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion/fnf_{1}:PackagingTypeFNF"/>
        <xsl:call-template name="for.item">
          <xsl:with-param name="item" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion"></xsl:with-param>
          <xsl:with-param name="biGTIN" select="$biGTINA"></xsl:with-param>
          <xsl:with-param name="biPackagingFNF" select="$biPackagingFNFA"></xsl:with-param>
          <xsl:with-param name="BI" select="$BA"></xsl:with-param>
        </xsl:call-template>
        <xsl:for-each select="fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion">
          <xsl:call-template name="for.item">
            <xsl:with-param name="item" select="."></xsl:with-param>
            <xsl:with-param name="biGTIN" select="$biGTINA"></xsl:with-param>
            <xsl:with-param name="biPackagingFNF" select="$biPackagingFNFA"></xsl:with-param>
            <xsl:with-param name="BI" select="$BA"></xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:if>
      <s:Row>
      </s:Row>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
