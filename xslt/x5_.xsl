<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
     xmlns:sinfos="http://schemas.sinfos.de/TradeItemMessages/1.2.0/TradeItemMessage"
				xmlns:fnf_{1}="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_{2}"
                xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
>
  <xsl:output method="xml" indent="yes"/>
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
          <xsl:value-of select="concat(substring($styleID,1,1),'2',substring($styleID,3))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$item/@status='added' or $node/@status='added'">
        <xsl:attribute name="s:StyleID">
          <xsl:value-of select="concat(substring($styleID,1,1),'3',substring($styleID,3))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$item/@status='deleted' or $node/@status='deleted'">
        <xsl:attribute name="s:StyleID">
          <xsl:value-of select="concat(substring($styleID,1,1),'4',substring($styleID,3))"/>
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
    <xsl:param name="BI"/>
    <s:Row s:AutoFitHeight="0" s:Height="62.25">
      <xsl:variable name="status" select="$item/@status"/>
      <xsl:choose>
        <xsl:when test="$status='changed'">
          <xsl:attribute name="ss:StyleID">s2011</xsl:attribute>
        </xsl:when>
        <xsl:when test="$status='deleted'">
          <xsl:attribute name="ss:StyleID">s4011</xsl:attribute>
        </xsl:when>
        <xsl:when test="$status='added'">
          <xsl:attribute name="ss:StyleID">s3011</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="ss:StyleID">s1011</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number"></s:Data>
          <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1004</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:BrandNameML/fnf_{1}:BrandNameML[@language='RU']"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1012</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$BI/fnf_{1}:BrandNameML/fnf_{1}:BrandNameML[@language='RU']"/>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:ItemNameLongML/fnf_{1}:ItemNameLongML[@language='RU']"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1002</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:ItemNameLongML/fnf_{1}:ItemNameLongML[@language='RU']"/>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:ContentUnitOfMeasure"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:choose>
            <xsl:when test="$item/fnf_{1}:ContentUnitOfMeasure">
              <xsl:value-of select="$item/fnf_{1}:Content"/>
              <xsl:text>&#160;</xsl:text>
              <mapTo list="measurement" firstcell="code" secondcell="short_rus">
                <xsl:value-of select="$item/fnf_{1}:ContentUnitOfMeasure"/>
              </mapTo>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$BI/fnf_{1}:Content"/>
              <xsl:text>&#160;</xsl:text>
              <mapTo list="measurement" firstcell="code" secondcell="short_rus">
                <xsl:value-of select="$BI/fnf_{1}:ContentUnitOfMeasure"/>
              </mapTo>
            </xsl:otherwise>
          </xsl:choose>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1005</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">шт.</s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:GTIN"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1006</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$BI/fnf_{1}:GTIN"/>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:InternalItemIDofSupplier"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:InternalItemIDofSupplier"/>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1013</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1013</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:NumberOfBaseUnitContained"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:NumberOfBaseUnitContained"/>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:PalletDataPalletLoadingHeight"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$item/fnf_{1}:PalletDataPalletLoadingHeight and $item/fnf_{1}:NumberOfLayersPerPallet">
          <xsl:variable name="sub">
            <xsl:value-of select="$item/fnf_{1}:PalletDataPalletLoadingHeight - 150"/>
          </xsl:variable>
          <xsl:value-of select="format-number($sub div $item/fnf_{1}:NumberOfLayersPerPallet div 1000,'#.00')"/>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1007</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet and $item/fnf_{1}:NumberOfLayersPerPallet">
            <xsl:value-of select="($item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet*$item/fnf_{1}:NumberOfBaseUnitContained) div $item/fnf_{1}:NumberOfLayersPerPallet"/>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:GrossWeightFNFTSS"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1008</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$BI/fnf_{1}:GrossWeightFNFTSS">
            <mapTo list="conversion" tocode="KGM" firstcell="code" secondcell="tocode" thirdcell="coefficient">
              <xsl:value-of select="$BI/fnf_{1}:GrossWeightUnitOfMeasure"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of select="$BI/fnf_{1}:GrossWeightFNFTSS"/>
            </mapTo>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:NetWeight"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$BI/fnf_{1}:NetWeight">
            <mapTo list="conversion" tocode="KGM" firstcell="code" secondcell="tocode" thirdcell="coefficient">
              <xsl:value-of select="$BI/fnf_{1}:NetWeightUnitOfMeasure"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of select="$BI/fnf_{1}:NetWeight"/>
            </mapTo>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MeasurementTradeItemLengthDepth"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1009</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$BI/fnf_{1}:MeasurementTradeItemLengthDepth">
            <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
              <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemHeightUnitOfMeasure"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemLengthDepth"/>
            </mapTo>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MeasurementTradeItemWidth"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1009</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$BI/fnf_{1}:MeasurementTradeItemWidth">
            <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
              <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemWidthUnitOfMeasure"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemWidth"/>
            </mapTo>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MeasurementTradeItemHeight"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1009</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$BI/fnf_{1}:MeasurementTradeItemHeight">
            <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
              <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemHeightUnitOfMeasure"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemHeight"/>
            </mapTo>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:ManufacturersName"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$BI/fnf_{1}:ManufacturersName"/>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:CountryOfOrigin"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1010</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:if test="$BI/fnf_{1}:CountryOfOrigin!='RU'">
            <xsl:text>имп.</xsl:text>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:CountryOfOrigin"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1010</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:if test="$BI/fnf_{1}:CountryOfOrigin='RU'">
            <xsl:text>отеч.</xsl:text>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item/fnf_{1}:CertificatesOfQuality"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1007</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:CertificatesOfQuality/fnf_{1}:certificatesOfQuality/fnf_{1}:certificateOfQuality/fnf_{1}:certificateOfQualityNumber"/>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1013</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$item"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:CountryOfOrigin"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1010</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <mapTo list="countries" firstcell="code" secondcell="rus"><xsl:value-of select="$BI/fnf_{1}:CountryOfOrigin"/></mapTo>
         </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:VAT_{0}"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:if test="$BI/fnf_{1}:VAT_{0}='59'">
            <xsl:text>s10</xsl:text>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:VAT_{0}"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String">
          <xsl:if test="$BI/fnf_{1}:VAT_{0}='60'">
            <xsl:text>18</xsl:text>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="Number">
          <xsl:if test="$BI/fnf_{1}:MinimumDurabilityFromArrival">
            <mapTo list="conversion" tocode="804" firstcell="code" secondcell="tocode" thirdcell="coefficient">
              <xsl:value-of select="$BI/fnf_{1}:MinimumDurabilityFromArrivalTimeUnit"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
            </mapTo>
          </xsl:if>
        </s:Data>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:NamedCell s:Name="_FilterDatabase"/>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String"></s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <s:Data s:Type="String"></s:Data>
      </s:Cell>
      <s:Cell>
        <xsl:call-template name="markcellchanges">
          <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
          <xsl:with-param name="item" select="$item"/>
          <xsl:with-param name="styleID">
            <xsl:text>s1003</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </s:Cell>
    </s:Row>
  </xsl:template>
  <xsl:template name="getbranches">
    <xsl:param name="parentGTIN" />
    <xsl:param name="BI" />
    <xsl:choose>
      <xsl:when test="//fnf_{1}:Item/fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTIN=$parentGTIN]/fnf_{1}:GTINofNextLowerPackagingItem">
        <xsl:choose>
          <xsl:when test="count(//fnf_{1}:Item/fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTINofNextLowerPackagingItem=$parentGTIN])=0">
            <xsl:call-template name="for.item">
              <xsl:with-param name="item" select="//fnf_{1}:Item/fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTIN=$parentGTIN]"/>
              <xsl:with-param name="BI" select="$BI"></xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:for-each select="//fnf_{1}:Item/fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTINofNextLowerPackagingItem=$parentGTIN]">
              <xsl:call-template name="getbranches">
                <xsl:with-param name="parentGTIN">
                  <xsl:value-of select="fnf_{1}:GTIN"/>
                </xsl:with-param>
                <xsl:with-param name="BI" select="$BI"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="for.item">
          <xsl:with-param name="item" select="//fnf_{1}:Item/fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTIN=$parentGTIN]"/>
          <xsl:with-param name="BI" select="$BI"></xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="/">
    <xsl:for-each select="fnf_{1}:Item">
      <xsl:variable name="biGTIN" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion/fnf_{1}:GTIN"/>
      <xsl:variable name="biGTINA" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion/fnf_{1}:GTIN"/>
      <xsl:variable name="BI" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion" />
      <xsl:variable name="BA" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion" />
      <xsl:if test="$biGTIN">
        <xsl:for-each select="fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTINofNextLowerPackagingItem=$BI/fnf_{1}:GTIN]">
          <xsl:call-template name="getbranches">
            <xsl:with-param name="parentGTIN">
              <xsl:value-of select="fnf_{1}:GTIN"/>
            </xsl:with-param>
            <xsl:with-param name="BI" select="$BI"></xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="$biGTINA">
        <xsl:for-each select="fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTINofNextLowerPackagingItem=$BA/fnf_{1}:GTIN]">
          <xsl:call-template name="getbranches">
            <xsl:with-param name="parentGTIN">
              <xsl:value-of select="fnf_{1}:GTIN"/>
            </xsl:with-param>
            <xsl:with-param name="BI" select="$BA"></xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:if>
      <s:Row>
      </s:Row>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
