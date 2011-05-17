<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
  <xsl:output method="xml" indent="yes"/>
  <xsl:template match="/">
    <?mso-application progid="Excel.Sheet"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"
        xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:sinfos="http://schemas.sinfos.de/TradeItemMessages/1.2.0/TradeItemMessage"
        xmlns:fnf_fnd_at="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_AT"
        xmlns:fnf_fnd_be="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_BE"
        xmlns:fnf_fnd_ch="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_CH"
        xmlns:fnf_fnd_cz="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_CZ"
        xmlns:fnf_fnd_de="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_DE"
       xmlns:fnf_fnd_dk="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_DK"
       xmlns:fnf_fnd_ee="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_EE"
       xmlns:fnf_fnd_es="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_ES"
       xmlns:fnf_fnd_fi="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_FI"
       xmlns:fnf_fnd_fr="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_FR"
       xmlns:fnf_fnd_gb="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_GB"
       xmlns:fnf_fnd_hu="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_HU"
       xmlns:fnf_fnd_ie="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_IE"
       xmlns:fnf_fnd_it="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_IT"
       xmlns:fnf_fnd_nl="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_NL"
       xmlns:fnf_fnd_pl="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_PL"
       xmlns:fnf_fnd_pt="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_PT"
       xmlns:fnf_fnd_ro="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_RO"
       xmlns:fnf_fnd_ru="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_RU"
       xmlns:fnf_fnd_se="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_SE"
       xmlns:fnf_fnd_ua="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_UA"
       xmlns:fnf_rap_at="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_AT"
       xmlns:fnf_rap_de="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_DE"
       xmlns:fnf_rap_dk="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_DK"
       xmlns:fnf_rap_ee="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_EE"
       xmlns:fnf_rap_fi="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_FI"
       xmlns:fnf_rap_pl="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_PL"
       xmlns:fnf_rap_ru="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_RU"
        xmlns:html="http://www.w3.org/TR/REC-html40">
      <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
        <LastAuthor>Rogozhin Anton Yurevich</LastAuthor>
        <Created>2009-11-22T10:48:42Z</Created>
        <Version>11.9999</Version>
      </DocumentProperties>
      <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
        <WindowHeight>10005</WindowHeight>
        <WindowWidth>10005</WindowWidth>
        <WindowTopX>120</WindowTopX>
        <WindowTopY>135</WindowTopY>
        <ProtectStructure>False</ProtectStructure>
        <ProtectWindows>False</ProtectWindows>
      </ExcelWorkbook>
      <Styles>
        <Style ss:ID="Default" ss:Name="Normal">
          <Alignment ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style ss:ID="s22">
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="s23">
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="s24">
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
        </Style>
      </Styles>
      <Worksheet ss:Name="Items">
        <Table ss:ExpandedColumnCount="71" x:FullColumns="1"
         x:FullRows="1">
          <Column ss:Width="99.75" ss:Span="70"/>
          <Row>
            <Cell>
              <Data ss:Type="String">Hierarchy</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">GTIN</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">StartValidityDate</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">GLNofDataSupplier</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">NameOfDataSupplier</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">TargetMarketCountryCode</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">SectorCode</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">BaseUnit</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ActionRequest</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">NonPublicTradeItem</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">NumberOfBaseUnitContained</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ItemNameLongML(Profile country)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ItemNameLongML(EN)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ItemName(Profile country)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ItemName(EN)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">DespatchUnit</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">InvoiceUnit</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">OrderingUnit</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">AdditionalInformation</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">BulkProduct</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ConsumerUnit</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">GLNofManufacturer</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ManufacturersName</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">Content</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ContentUnitOfMeasure</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">GrossWeightFNFTSS (GRM)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">NetWeight (KGM) </Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">BasePriceDeclarationRelevant</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">VAT</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">BrandNameML(Profile country)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">BrandNameML(EN)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">SubBrand</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">PLUDescriptionML(Profile country)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">PLUDescriptionML(EN)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">GPC</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">PricingOnTheProduct</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">freshnessDateOnProduct</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">RightOfReturnForNonSoldTradeItems</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">DangerousGoodsIndication</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">DangerousSubstancesDeclaration</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">MaterialSafetyDataSheet</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">Biocide</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">CountryOfOrigin</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">BatchNumbered</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">PackagingTypeFNF</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">MeasurementTradeItemHeight (MMT)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">MeasurementTradeItemLengthDepth (MMT)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">MeasurementTradeItemWidth (MMT)</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">Barcoded</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">TemperatureMinStorage</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">TemperatureMaxStorage</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">TemperatureUnitOfMeasure</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">TradeItemIsReleased</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ReleasedAtDateTime</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">LastChangedDateTime</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">CreatedAtDateTime</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ChangedBy</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">GlobalTradeItemIndicator</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">ReleaseStatusOfTradeItem</Data>
            </Cell>
            <Cell>
              <Data ss:Type="String">GLOBAL_ActionRequest</Data>
            </Cell>
          </Row>
          <root />
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
          <PageSetup>
            <PageMargins x:Bottom="0.984251969" x:Left="0.78740157499999996"
             x:Right="0.78740157499999996" x:Top="0.984251969"/>
          </PageSetup>
          <Selected/>
          <Panes>
            <Pane>
              <Number>3</Number>
              <ActiveRow>14</ActiveRow>
              <ActiveCol>1</ActiveCol>
            </Pane>
          </Panes>
          <ProtectObjects>False</ProtectObjects>
          <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
      </Worksheet>
    </Workbook>
      </xsl:template>
</xsl:stylesheet>
