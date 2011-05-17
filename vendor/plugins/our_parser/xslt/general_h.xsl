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
    <xsl:element name="s:Cell">
    <xsl:choose>
      <xsl:when test="$node/@status='changed'">
        <xsl:attribute name="s:StyleID">s22</xsl:attribute>
      </xsl:when>
      <xsl:when test="$node/@status='added'">
        <xsl:attribute name="s:StyleID">s24</xsl:attribute>
      </xsl:when>
      <xsl:when test="$node/@status='deleted'">
        <xsl:attribute name="s:StyleID">s23</xsl:attribute>
      </xsl:when>
    </xsl:choose>
      <s:Data s:Type="String">
        <xsl:value-of select="$node"/>
      </s:Data>
    </xsl:element>
  </xsl:template>
  <xsl:template name="for.item">
    <xsl:param name="firstColumn"/>
    <xsl:param name="item" />
    <xsl:element name="s:Row">
      <xsl:variable name="status" select="$item/@status"/>
      <xsl:choose>
        <xsl:when test="$status='changed'">
          <xsl:attribute name="ss:StyleID">s22</xsl:attribute>
        </xsl:when>
        <xsl:when test="$status='deleted'">
          <xsl:attribute name="ss:StyleID">s23</xsl:attribute>
        </xsl:when>
        <xsl:when test="$status='added'">
          <xsl:attribute name="ss:StyleID">s24</xsl:attribute>
        </xsl:when>
      </xsl:choose>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$firstColumn"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:GTIN"/>
        </s:Data>
      </s:Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:StartValidityDate"/>
      </xsl:call-template>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:GLNofDataSupplier"/>
        </s:Data>
      </s:Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:NameOfDataSupplier"/>
      </xsl:call-template>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:TargetMarketCountryCode"/>
        </s:Data>
      </s:Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:SectorCode"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:BaseUnit"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ActionRequest"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:NonPublicTradeItem"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:NumberOfBaseUnitContained"/>
      </xsl:call-template>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:ItemNameLongML/fnf_{1}:ItemNameLongML[@language='{0}']"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:ItemNameLongML/fnf_{1}:ItemNameLongML[@language='EN']"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='{0}']"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:ItemName/fnf_{1}:ItemName[@language='EN']"/>
        </s:Data>
      </s:Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:DespatchUnit"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:InvoiceUnit"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:OrderingUnit"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:AdditionalInformation"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:BulkProduct"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ConsumerUnit"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:GLNofManufacturer"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ManufacturersName"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:Content"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ContentUnitOfMeasure"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:GrossWeightFNFTSS"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:NetWeight"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:BasePriceDeclarationRelevant"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:VAT_{0}"/>
      </xsl:call-template>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:BrandNameML/fnf_{1}:BrandNameML[@language='{0}']"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:BrandNameML/fnf_{1}:BrandNameML[@language='EN']"/>
        </s:Data>
      </s:Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:SubBrand"/>
      </xsl:call-template>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:PLUDescriptionML/fnf_{1}:PLUDescriptionML[@language='{0}']"/>
        </s:Data>
      </s:Cell>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:PLUDescriptionML/fnf_{1}:PLUDescriptionML[@language='EN']"/>
        </s:Data>
      </s:Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:GPC"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:PricingOnTheProduct"/>
      </xsl:call-template>
      <s:Cell>
        <s:Data s:Type="String">
          <xsl:value-of select="$item/fnf_{1}:freshnessDateOnProduct/fnf_{1}:freshnessDatesOnProduct/fnf_{1}:freshnessDateOnProduct/fnf_{1}:freshnessDateOnProduct"/>
        </s:Data>
      </s:Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:RightOfReturnForNonSoldTradeItems"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:DangerousGoodsIndication"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:DangerousSubstancesDeclaration"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:MaterialSafetyDataSheet"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:Biocide"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:DangerousSubstancesDeclaration"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:CountryOfOrigin"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:BatchNumbered"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:PackagingTypeFNF"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemHeight"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemLengthDepth"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemWidth"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:Barcoded"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:TemperatureMinStorage"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:TemperatureMaxStorage"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:TemperatureUnitOfMeasure"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:TradeItemIsReleased"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ReleasedAtDateTime"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:LastChangedDateTime"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:CreatedAtDateTime"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ChangedBy"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:GlobalTradeItemIndicator"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ReleaseStatusOfTradeItem"/>
      </xsl:call-template>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:GLOBAL_ActionRequest"/>
      </xsl:call-template>
      </xsl:element>
  </xsl:template>
  <xsl:template match="/">
    <xsl:for-each select="fnf_{1}:Item">
      <xsl:variable name="GTIN" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion/fnf_{1}:GTIN"/>
      <xsl:variable name="GTINA" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion/fnf_{1}:GTIN"/>
      <xsl:variable name="Item" select="."/>
      <xsl:if test="$GTIN">
        <xsl:for-each select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion">
          {3}
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="$GTINA">
        <xsl:for-each select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion">
          {4}
        </xsl:for-each>
      </xsl:if>
      <s:Row>
      </s:Row>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>