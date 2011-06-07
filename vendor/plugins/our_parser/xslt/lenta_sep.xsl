<?xml version="1.0"  encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
     xmlns:sinfos="http://schemas.sinfos.de/TradeItemMessages/1.2.0/TradeItemMessage"
				xmlns:fnf_{1}="http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_{2}"
                xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"
>
  <xsl:output method="xml" indent="yes"/>
  <xsl:template name="markcellchanges">
    <xsl:param name="node"/>
    <xsl:param name="styleID"/>
    <xsl:choose>
      <xsl:when test="$node/@status='changed'">
        <xsl:attribute name="s:StyleID">
          <xsl:value-of select="concat(substring($styleID,1,1),'2',substring($styleID,3))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$node/@status='added'">
        <xsl:attribute name="s:StyleID">
          <xsl:value-of select="concat(substring($styleID,1,1),'3',substring($styleID,3))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$node/@status='deleted'">
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
  <xsl:template name="dataout">
    <xsl:param name="item" />
    <xsl:param name="BI"/> 
   <?mso-application progid="Excel.Sheet"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
     xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel"
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
  <Author>dana.yakovleva</Author>
  <LastAuthor>Rogozhin Anton Yurevich</LastAuthor>
  <LastPrinted>2008-11-01T11:13:12Z</LastPrinted>
  <Created>2006-11-17T11:49:41Z</Created>
  <LastSaved>2009-10-12T12:51:40Z</LastSaved>
  <Company>Lenta Cash&amp;Carry</Company>
  <Version>11.9999</Version>
 </DocumentProperties>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>7620</WindowHeight>
  <WindowWidth>12120</WindowWidth>
  <WindowTopX>360</WindowTopX>
  <WindowTopY>240</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s18" ss:Name="Денежный">
   <NumberFormat
    ss:Format="_-* #,##0.00\ &quot;Kč&quot;_-;\-* #,##0.00\ &quot;Kč&quot;_-;_-* &quot;-&quot;??\ &quot;Kč&quot;_-;_-@_-"/>
  </Style>
  <Style ss:ID="s16" ss:Name="Финансовый">
   <NumberFormat
    ss:Format="_-* #,##0.00\ _K_č_-;\-* #,##0.00\ _K_č_-;_-* &quot;-&quot;??\ _K_č_-;_-@_-"/>
  </Style>
  <Style ss:ID="m20674128">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20674138">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <NumberFormat ss:Format="Short Date"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673976">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673986">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673996">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20674006">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673824">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673834">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673844">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673854">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673672">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673682">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16" ss:Bold="1"/>
  </Style>
  <Style ss:ID="m20673692">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20673702">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20672338">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20672358">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20672368">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20672378">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20672388">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
  </Style>
  <Style ss:ID="m20671176">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20671186">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20671196">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat ss:Format="Fixed"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20671018">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20671038">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat ss:Format="Fixed"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670856">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670866">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670876">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670704">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670714">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670724">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Color="#FF0000" ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670744">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670754">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670552">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670562">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670572">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20678254">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20678274">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670430">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670248">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20670278">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20669792">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20669802">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20669812">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20669822">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668826">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668836">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668846">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668664">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668674">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668684">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668694">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20663752">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20663762">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20663772">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20663782">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668522">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668532">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668542">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668360">
   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668370" ss:Parent="s18">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668380">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668390">
   <Alignment ss:Horizontal="Right" ss:Vertical="Center" ss:Indent="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20668400">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20663620">
   <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <NumberFormat/>
  </Style>
  <Style ss:ID="m20663510">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20657504">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Interior/>
   <NumberFormat ss:Format="0%"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20657524">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <NumberFormat ss:Format="0%"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662168">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662188">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662198">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662218">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662016">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662036">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662046">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20662066">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20646452">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="18"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20646472">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <NumberFormat ss:Format="Short Date"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20646482">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="m20646502">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s21">
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s22">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s23">
   <Font ss:FontName="Arial CE" x:CharSet="238"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s24">
   <Alignment ss:Vertical="Bottom"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s25">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s26">
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s27">
   <Alignment ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s28">
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s29">
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s30">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s31">
   <Alignment ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s33">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="20" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s34">
   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="20" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s35">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s36">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s39" ss:Parent="s16">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s40">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s41">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"
    ss:Italic="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s43">
   <Alignment ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
  </Style>
  <Style ss:ID="s44">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s46">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s47">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s48">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s49">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s50">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <Protection x:HideFormula="1"/>
  </Style>
  <Style ss:ID="s51">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s55">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s57">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s73">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s81">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s87">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s95">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s105">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s110">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s114">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s119">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <NumberFormat ss:Format="0%"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s125" ss:Parent="s18">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s131" ss:Parent="s18">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s132" ss:Parent="s18">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s134">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s138">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s139">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s141" ss:Parent="s18">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s143">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s144">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s145">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <NumberFormat ss:Format="0%"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s147">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s151">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <NumberFormat ss:Format="0"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s152">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s153">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s159">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s161">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s162">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s163">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s164" ss:Parent="s18">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s165" ss:Parent="s18">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s166" ss:Parent="s18">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s178" ss:Parent="s18">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s184">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s186">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s192">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s196">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s197">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s198">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s199">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s200">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s201">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:Indent="3"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s202">
   <Alignment ss:Horizontal="Right" ss:Vertical="Center" ss:Indent="2"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s219">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s220">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="8"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s221">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s244">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s265">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s274">
   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s279">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s285">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s286">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s287">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s288">
   <Alignment ss:Horizontal="Center" ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s289">
   <Alignment ss:Horizontal="Center" ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s290">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s291">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="12"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s292">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s293">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s294">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s295">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="12"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s297">
   <Alignment ss:Horizontal="Left" ss:Vertical="Justify" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s305">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s306">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s307">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="12"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s308">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="12"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s309">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s310">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s312">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s316">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s317">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s318">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s319">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s328">
   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s332">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s333">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s334">
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s335">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s336">
   <Alignment ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s337">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s338">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="12"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s346">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16" ss:Italic="1"/>
  </Style>
  <Style ss:ID="s347">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s348">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s349">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s350">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s351">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s360">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s380">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s392">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s393">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="12"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s394">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s395">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s402">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s410">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s411">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s414">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s418">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s424">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s426">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s441">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s442">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s443">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s444">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s445">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s446">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s454">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s457">
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s458">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s460">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s461">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="Short Date"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s462">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="Short Date"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s463">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Color="#000000" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s464">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16" ss:Color="#000000"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s465">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s466">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Color="#FF0000" ss:Bold="1"/>
   <NumberFormat ss:Format="Short Date"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s467">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Color="#FF0000" ss:Bold="1"/>
   <NumberFormat ss:Format="Short Date"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s468">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Color="#FF0000" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s470">
   <Alignment ss:Horizontal="Center" ss:Vertical="Justify"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s474">
   <Alignment ss:Vertical="Justify"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s475">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s476">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s477">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s478">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s484">
   <Alignment ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s485">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s486">
   <Alignment ss:Horizontal="Left" ss:Vertical="Top"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Color="#FFFFFF" ss:Bold="1" ss:Italic="1"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s487">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1" ss:Italic="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s488">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s489">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s490">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s491">
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s492">
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s493">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s494">
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s495">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s496">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s497">
   <Borders>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s498">
   <Borders>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s501">
   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
   <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s530">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s531">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s532">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s533">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s534">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s535">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s536">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s537">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s556">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s565">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16" ss:Bold="1"/>
  </Style>
  <Style ss:ID="s568">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s575">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s576">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <NumberFormat ss:Format="Short Date"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s577">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s578">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="204" ss:Size="16" ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s579">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s580">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s581">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s582">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
  </Style>
  <Style ss:ID="s583">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
  </Style>
  <Style ss:ID="s584">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
   </Borders>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Gray125" ss:PatternColor="#FFFFFF"/>
  </Style>
  <Style ss:ID="s585">
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s593">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s600">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
  </Style>
  <Style ss:ID="s602">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Left" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="14"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s606">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s607">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s608">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Right" ss:LineStyle="Double" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s609">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders/>
   <Font x:CharSet="204" x:Family="Swiss" ss:Size="16"/>
   <Interior/>
   <Protection ss:Protected="0"/>
  </Style>
   <Style ss:ID="s1001">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="18"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1002">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1003">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior/>
     <NumberFormat ss:Format="0%"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1004" ss:Parent="s18">
     <Alignment ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <NumberFormat/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1005" ss:Parent="s18">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <NumberFormat/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1006">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1007">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Left" ss:LineStyle="Continuous"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1008">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1009">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1010">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1011">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1012">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s1013">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
     <Protection ss:Protected="0"/>
   </Style>
  <Style ss:ID="s2001">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="18"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2002">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2003">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0%"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2004" ss:Parent="s18">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2005" ss:Parent="s18">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <NumberFormat/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2006">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2007">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2008">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
   <Style ss:ID="s2009">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2010">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Protection ss:Protected="0"/>
  </Style>
 <Style ss:ID="s2011">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2012">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
    ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0"/>
   <Protection ss:Protected="0"/>
  </Style>
  <Style ss:ID="s2013">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
   <Protection ss:Protected="0"/>
  </Style>
   <Style ss:ID="s3001">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="18"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3002">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3003">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0%"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3004" ss:Parent="s18">
     <Alignment ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <NumberFormat/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3005" ss:Parent="s18">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <NumberFormat/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3006">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3007">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Left" ss:LineStyle="Continuous"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3008">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3009">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3010">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3011">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3012">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s3013">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4001">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="18"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4002">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4003">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0%"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4004" ss:Parent="s18">
     <Alignment ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <NumberFormat/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4005" ss:Parent="s18">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <NumberFormat/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4006">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4007">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Left" ss:LineStyle="Continuous"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4008">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4009">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4010">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"
      ss:Bold="1"/>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4011">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4012">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="14"
      ss:Bold="1"/>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <NumberFormat ss:Format="0"/>
     <Protection ss:Protected="0"/>
   </Style>
   <Style ss:ID="s4013">
     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
     <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
     <Font ss:FontName="Arial CE" x:CharSet="238" x:Family="Swiss" ss:Size="16"/>
     <Protection ss:Protected="0"/>
   </Style>
 </Styles>
 <Worksheet ss:Name="Лист2">
  <Table ss:ExpandedColumnCount="13" ss:ExpandedRowCount="81" x:FullColumns="1"
   x:FullRows="1" ss:StyleID="s21">
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="212.25"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="93"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="109.5"/>
   <Column ss:StyleID="s22" ss:AutoFitWidth="0" ss:Width="72"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="92.25"/>
   <Column ss:StyleID="s23" ss:AutoFitWidth="0" ss:Width="37.5"/>
   <Column ss:StyleID="s22" ss:AutoFitWidth="0" ss:Width="88.5"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="99"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="78.75"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="93.75"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="99"/>
   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="107.25"/>
   <Column ss:StyleID="s24" ss:AutoFitWidth="0" ss:Width="47.25"/>
   <Row ss:AutoFitHeight="0" ss:Height="20.25" ss:StyleID="s26">
    <Cell ss:StyleID="s28"><Data ss:Type="String">Приложение № 4 к договору</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s25"/>
    <Cell ss:Index="7" ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="20.25" ss:StyleID="s26">
    <Cell ss:Index="2" ss:StyleID="s28"/>
    <Cell ss:StyleID="s28"/>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="7" ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="18.75" ss:StyleID="s29">
    <Cell ss:Index="2" ss:MergeAcross="7" ss:StyleID="s33"><Data ss:Type="String">Карта ввода (уведомление о логистических параметрах)</Data></Cell>
    <Cell ss:StyleID="s34"/>
    <Cell ss:StyleID="s34"/>
    <Cell ss:StyleID="s34"/>
    <Cell ss:StyleID="s34"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="37.5" ss:StyleID="s29">
    <Cell ss:Index="2" ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s36"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
    <Cell ss:StyleID="s35"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="57" ss:StyleID="s26">
    <Cell ss:MergeAcross="4" ss:StyleID="s39"><Data ss:Type="String">Области выделенные полужирным шрифтом заполняются поставщиком</Data></Cell>
    <Cell ss:StyleID="s40"/>
    <Cell ss:MergeAcross="5" ss:StyleID="s41"><Data ss:Type="String">Заполняется в Ленте</Data></Cell>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="20.25" ss:StyleID="s26">
    <Cell ss:StyleID="s43"/>
    <Cell ss:StyleID="s43"/>
    <Cell ss:StyleID="s43"/>
    <Cell ss:StyleID="s43"/>
    <Cell ss:StyleID="s43"/>
    <Cell ss:StyleID="s40"/>
    <Cell ss:StyleID="s44"/>
    <Cell ss:StyleID="s41"/>
    <Cell ss:StyleID="s41"/>
    <Cell ss:StyleID="s41"/>
    <Cell ss:StyleID="s41"/>
    <Cell ss:StyleID="s41"/>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:MergeAcross="4" ss:StyleID="s46"><Data ss:Type="String">ОСНОВНЫЕ ДАННЫЕ ПРОДУКТА</Data></Cell>
    <Cell ss:StyleID="s48"/>
    <Cell ss:StyleID="s25"/>
    <Cell ss:StyleID="s49"/>
    <Cell ss:StyleID="s49"/>
    <Cell ss:StyleID="s50"/>
    <Cell ss:StyleID="s50"/>
    <Cell ss:StyleID="s50"/>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="72.75" ss:StyleID="s26">
    <Cell ss:StyleID="s51"><Data ss:Type="String">Название товара: </Data></Cell>
    <Cell ss:MergeAcross="3">
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:ItemNameLongML/fnf_{1}:ItemNameLongML[@language='RU']"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1001</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
     <Data ss:Type="String">
         <xsl:value-of select="$item/fnf_{1}:ItemNameLongML/fnf_{1}:ItemNameLongML[@language='RU']"/>
          <!--<xsl:if test="$BI/fnf_{1}:ContentUnitOfMeasure">
            <xsl:text>&#160;</xsl:text>
            <xsl:value-of select="$BI/fnf_{1}:Content"/>
            <xsl:text>&#160;</xsl:text>
            <mapTo list="measurement" firstcell="code" secondcell="short_rus">
              <xsl:value-of select="$BI/fnf_{1}:ContentUnitOfMeasure"/>
            </mapTo>
          </xsl:if>-->
       </Data>
     </Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П1</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s57"><Data ss:Type="String">№ товара SAP:</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20646472"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">О1</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="40.5" ss:StyleID="s26">
    <Cell ss:MergeAcross="4" ss:StyleID="m20646482"/>
    <Cell ss:StyleID="s55"/>
    <Cell ss:MergeAcross="2" ss:StyleID="s73"><Data ss:Type="String">Дерево каталогов (код подкатегории)</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20646502"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ1</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="35.25" ss:StyleID="s26">
    <Cell ss:StyleID="s81"><Data ss:Type="String">Название поставщика: </Data></Cell>
    <Cell ss:MergeAcross="3">
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$BI/fnf_{1}:NameOfDataSupplier"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1002</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="String">
        <xsl:value-of select="$BI/fnf_{1}:NameOfDataSupplier"/>
      </Data>
    </Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П2</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s87"><Data ss:Type="String">Номер поставщика SAP</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20662036"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ2</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="35.25" ss:StyleID="s26">
    <Cell ss:StyleID="s95"><Data ss:Type="String">Контактное лицо: </Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20662046"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П3</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s73"><Data ss:Type="String">Группа закупок (направление КС):</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20662066"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ3</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="30.75" ss:StyleID="s26">
    <Cell ss:StyleID="s95"><Data ss:Type="String">Тел/факс:  </Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20662168"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П4</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s105"><Data ss:Type="String">Группа закупок (секция):</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20662188"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ4</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="30.75" ss:StyleID="s26">
    <Cell ss:StyleID="s95"><Data ss:Type="String">ИНН поставщика: </Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20662198"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П5</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s110"><Data ss:Type="String">Вид товара:</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20662218"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ5</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="30.75" ss:StyleID="s26">
    <Cell ss:StyleID="s114"><Data ss:Type="String">Ставка / КОД НДС закупка:  </Data><Comment
      ss:Author="vladimir.gorichev"><ss:Data
       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"
         x:CharSet="238" html:Size="8" html:Color="#000000">vladimir.gorichev:</Font></B><Font
        html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#000000">&#10;Нам важнее код НДС, который определяет ставку</Font></ss:Data></Comment></Cell>
    <Cell ss:MergeAcross="3">
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$BI/fnf_{1}:VAT_{0}"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1003</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$BI/fnf_{1}:VAT_{0}='59'">
          <xsl:text>0.10</xsl:text>
        </xsl:if>
       <xsl:if test="$BI/fnf_{1}:VAT_{0}='60'">
          <xsl:text>0.18</xsl:text>
        </xsl:if>
      </Data>
     </Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П6</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s73"><Data ss:Type="String">Ставка / КОД НДС продажа:</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20657524"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ6</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="54" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="s125"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>Название алкогольного производителя<Font
        html:Size="14"> (для импортного производства - название импортера):</Font></B></ss:Data></Cell>
    <Cell ss:StyleID="s131"/>
    <Cell ss:StyleID="s131"/>
    <Cell ss:StyleID="s132"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П7</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="s134"><Data ss:Type="String">№ алкогольного производителя SAP</Data></Cell>
    <Cell ss:StyleID="s138"/>
    <Cell ss:StyleID="s139"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ7</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="36.75" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="s141"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>ИНН производителя алкоголя <Font
        html:Size="14">(для импортного производства - ИНН импортера):</Font></B></ss:Data></Cell>
    <Cell ss:StyleID="s131"/>
    <Cell ss:StyleID="s131"/>
    <Cell ss:StyleID="s132"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П8</Data></Cell>
    <Cell ss:StyleID="s143"><Data ss:Type="String">Тип товара:</Data></Cell>
    <Cell ss:StyleID="s144"/>
    <Cell ss:StyleID="s144"/>
    <Cell ss:StyleID="s145"/>
    <Cell ss:StyleID="s145"/>
    <Cell ss:StyleID="s119"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ8</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="36.75" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="s141"><Data ss:Type="String">Субъект РФ производства алкоголя:</Data></Cell>
    <Cell ss:StyleID="s131"/>
    <Cell ss:StyleID="s131"/>
    <Cell ss:StyleID="s132"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П9</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s73"><Data ss:Type="String">Стандартная цена (производство):</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20663510"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ9</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.25" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="s147"><Data ss:Type="String">Код ОК 005 (ОКП) согласно сертификата соответствия</Data></Cell>
    <Cell ss:StyleID="s151"/>
    <Cell ss:StyleID="s152"/>
    <Cell ss:StyleID="s153"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П10</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s73"><Data ss:Type="String">Мин. остаточный срок годности:</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20663620"><Data ss:Type="String">час / сут / мес / лет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ10</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.25" ss:StyleID="s26">
    <Cell ss:StyleID="s95"><Data ss:Type="String">Продажа на вес</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s159"/>
    <Cell ss:StyleID="s161"><Data ss:Type="String">нет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П11</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s105"><Data ss:Type="String">Вид ассортиментного списка:</Data></Cell>
    <Cell ss:StyleID="s162"/>
    <Cell ss:StyleID="s162"/>
    <Cell ss:StyleID="s163"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ11</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.25" ss:StyleID="s26">
    <Cell ss:StyleID="s164"><Data ss:Type="String">Срок годности (гарантийный срок): </Data></Cell>
    <Cell ss:StyleID="s165"/>
    <Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1004</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <xsl:choose>
        <xsl:when test="$BI/fnf_{1}:MinimumDurabilityFromArrival">
         <Data ss:Type="Number">
              <mapTo list="conversion" tocode="804" firstcell="code" secondcell="tocode" thirdcell="coefficient">
                <xsl:value-of select="$BI/fnf_{1}:MinimumDurabilityFromArrivalTimeUnit"/>
                <xsl:text>:</xsl:text>
                <xsl:value-of select="$BI/fnf_{1}:MinimumDurabilityFromArrival"/>
              </mapTo>
          </Data>
        </xsl:when>
        <xsl:otherwise>
          <Data ss:Type="String"></Data>
        </xsl:otherwise>
      </xsl:choose>
    </Cell>
    <Cell ss:StyleID="s165"><Data ss:Type="String">дней</Data></Cell>
    <Cell ss:StyleID="s166"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П12</Data></Cell>
    <Cell ss:MergeAcross="5" ss:MergeDown="4" ss:StyleID="m20668360"><Data
      ss:Type="String">Дополнительные характеристики товара  (производство/NONFOOD):</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ12</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.25" ss:StyleID="s26">
    <Cell ss:StyleID="s178"><Data ss:Type="String">Страна происхождения:</Data></Cell>
    <Cell ss:MergeAcross="3">
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$BI/fnf_{1}:CountryOfOrigin"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1005</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="String">
         <mapTo list="countries" firstcell="code" secondcell="rus">
            <xsl:value-of select="$BI/fnf_{1}:CountryOfOrigin"/>
          </mapTo>
      </Data>
     </Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П13</Data></Cell>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.25" ss:StyleID="s26">
    <Cell ss:StyleID="s184"><Data ss:Type="String">Торговая марка (на языке оригинала)</Data></Cell>
    <Cell ss:MergeAcross="3">
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$BI/fnf_{1}:BrandNameML/fnf_{1}:BrandNameML[@language='RU']"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1006</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="String">
        <xsl:value-of select="$BI/fnf_{1}:BrandNameML/fnf_{1}:BrandNameML[@language='RU']"/>
      </Data>
    </Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П14</Data></Cell>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.25" ss:StyleID="s26">
    <Cell ss:StyleID="s184"><Data ss:Type="String">Сухой вес:</Data></Cell>
    <Cell ss:StyleID="s186"/>
    <Cell ss:MergeAcross="2" ss:StyleID="m20668390"><Data ss:Type="String">гр / мл</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П15</Data></Cell>
    <Cell ss:Index="13" ss:StyleID="s55"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.25" ss:StyleID="s26">
    <Cell ss:StyleID="s184"><Data ss:Type="String">Алкоголь (%)</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20668400"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П16</Data></Cell>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="32.25" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="s192"><Data ss:Type="String">Базовая единица измерения (БЕИ):</Data></Cell>
    <Cell ss:StyleID="s196"/>
    <Cell ss:StyleID="s196"/>
    <Cell ss:StyleID="s197"><Data ss:Type="String">Шт</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П17</Data></Cell>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s55"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="27" ss:StyleID="s26">
    <Cell ss:StyleID="s198"/>
    <Cell ss:StyleID="s199"/>
    <Cell ss:StyleID="s200"><Data ss:Type="String">Кол-во БЕИ</Data></Cell>
    <Cell ss:StyleID="s201"><Data ss:Type="String">Наименование</Data></Cell>
    <Cell ss:StyleID="s202"/>
    <Cell ss:Index="7" ss:MergeAcross="1" ss:StyleID="m20668522"><Data
      ss:Type="String">Сезон</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668532"><Data ss:Type="String">Коллекция</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668542"><Data ss:Type="String">Год</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ13</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="42" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="m20663752"><Data ss:Type="String">Дополнительная единица измерения</Data></Cell>
    <Cell ss:StyleID="s219"><Data ss:Type="String">короб</Data></Cell>
    <Cell ss:StyleID="s220"/>
    <Cell ss:StyleID="s221"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П18</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20663762"><Data ss:Type="String">СЗН</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20663772"><Data ss:Type="String">Весна-лето</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20663782"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ14</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="49.5" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="m20668664"><Data ss:Type="String">Дополнительная единица измерения</Data></Cell>
    <Cell ss:StyleID="s219"><Data ss:Type="String">паллета</Data></Cell>
    <Cell ss:StyleID="s220"/>
    <Cell ss:StyleID="s221"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П19</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668674"><Data ss:Type="String">СЗН</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668684"><Data ss:Type="String">Осень-зима</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668694"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ15</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="27" ss:StyleID="s26">
    <Cell ss:MergeAcross="4" ss:StyleID="s47"><Data ss:Type="String">ЛОГИСТИЧЕСКИЕ ДАННЫЕ</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П20</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668826"><Data ss:Type="String">СЗН</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668836"><Data ss:Type="String"
      x:Ticked="1">&#45;&#45;&#45;&#45;&#45;-</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20668846"><Data ss:Type="Number"></Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ16</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="39.75" ss:StyleID="s29">
    <Cell ss:StyleID="s244"><Data ss:Type="String" x:Ticked="1">Единица измерения (размещения) заказа (ЕИЗ):</Data><Comment
      ss:Author="ekaterina.zarayskaya"><ss:Data
       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"
         x:CharSet="238" html:Size="8" html:Color="#000000">ekaterina.zarayskaya:</Font></B><Font
        html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#000000">&#10;причем коробка - это ен обязательно мин.кол-во заказа</Font></ss:Data></Comment></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20669792"><Data ss:Type="Number"></Data><Comment
      ss:Author="luchal"><ss:Data xmlns="http://www.w3.org/TR/REC-html40"><B><Font
         html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#000000">figure of pieces in one carton</Font></B></ss:Data></Comment></Cell>
    <Cell ss:Index="7" ss:MergeAcross="1" ss:StyleID="m20669802"><Data
      ss:Type="String">&#45;&#45;&#45;&#45;&#45;-</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20669812"><Data ss:Type="String"
      x:Ticked="1">&#45;&#45;&#45;&#45;&#45;-</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20669822"><Data ss:Type="String"
      x:Ticked="1">&#45;&#45;&#45;&#45;&#45;&#45;-</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ17</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="41.25" ss:StyleID="s29">
    <Cell ss:StyleID="s265"><Data ss:Type="String">Минимальный заказ в ЕИЗ</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20670248"><Data ss:Type="String">1 короб</Data><Comment
      ss:Author="luchal"><ss:Data xmlns="http://www.w3.org/TR/REC-html40"><B><Font
         html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#000000">figure of cartons in one layer</Font></B></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П21</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s274"><Data ss:Type="String">Расчет задним числом</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s279"><Data ss:Type="String">да</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670278"><Data ss:Type="String">нет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ18</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="60" ss:StyleID="s29">
    <Cell ss:StyleID="s265"/>
    <Cell ss:StyleID="s285"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40">Шт.<Font html:Size="12">(не весовой товар)</Font></ss:Data></Cell>
    <Cell ss:StyleID="s286"><Data ss:Type="String">Трансп. упаковка</Data></Cell>
    <Cell ss:StyleID="s286"><Data ss:Type="String">Слой</Data></Cell>
    <Cell ss:StyleID="s287"><Data ss:Type="String">Паллет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П22</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s274"><Data ss:Type="String">Группа бренда:</Data></Cell>
    <Cell ss:StyleID="s288"/>
    <Cell ss:StyleID="s288"/>
    <Cell ss:StyleID="s288"><Data ss:Type="Number">2</Data></Cell>
    <Cell ss:StyleID="s289"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ19</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="42" ss:StyleID="s26">
    <Cell ss:StyleID="s290"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>Количество </B><Font>БЕИ (шт/весовой - гр)</Font></ss:Data></Cell>
    <Cell ss:StyleID="s291"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:NumberOfBaseUnitContained"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1007</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:value-of select="$item/fnf_{1}:NumberOfBaseUnitContained"/>
      </Data>
    </Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$item/fnf_{1}:NumberOfLayersPerPallet"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1007</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet and $item/fnf_{1}:NumberOfLayersPerPallet">
          <xsl:value-of select="($item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet*$item/fnf_{1}:NumberOfBaseUnitContained) div $item/fnf_{1}:NumberOfLayersPerPallet"/>
        </xsl:if>   
      </Data>
    </Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1008</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet and $item/fnf_{1}:NumberOfBaseUnitContained">
          <xsl:value-of select="$item/fnf_{1}:PalletDataNumberOfDespatchUnitsPerPallet*$item/fnf_{1}:NumberOfBaseUnitContained"/>
        </xsl:if>
      </Data>
    </Cell>
    <Cell ss:Index="7" ss:MergeAcross="1" ss:StyleID="s274"><Data ss:Type="String">Переменная ЕИ Заказа:</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s279"><Data ss:Type="String">да</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670430"><Data ss:Type="String">нет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ20</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="45" ss:StyleID="s26">
    <Cell ss:StyleID="s290"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>Длина                            </B><Font>мм.</Font></ss:Data></Cell>
    <Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$BI/fnf_{1}:MeasurementTradeItemLengthDepth"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1009</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$BI/fnf_{1}:MeasurementTradeItemLengthDepth">
          <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemLengthDepthUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemLengthDepth"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
    <Cell>
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemLengthDepth"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1009</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$item/fnf_{1}:MeasurementTradeItemLengthDepth">
          <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemLengthDepthUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemLengthDepth"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
    <Cell ss:StyleID="s291"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s295"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П23</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s297"><Data ss:Type="String">Планируемый оборот:</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20678254"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ21</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="47.25" ss:StyleID="s26">
    <Cell ss:StyleID="s290"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>Ширина                         </B><Font>мм.</Font></ss:Data></Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$BI/fnf_{1}:MeasurementTradeItemWidth"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1009</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$BI/fnf_{1}:MeasurementTradeItemWidth">
          <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemWidthUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemWidth"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
    <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemWidth"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1009</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$item/fnf_{1}:MeasurementTradeItemWidth">
          <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemWidthUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemWidth"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
    <Cell ss:StyleID="s291"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s295"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П24</Data></Cell>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21.75" ss:StyleID="s26">
    <Cell ss:StyleID="s290"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>Высота                          </B><Font>мм.</Font></ss:Data></Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$BI/fnf_{1}:MeasurementTradeItemHeight"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1009</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$BI/fnf_{1}:MeasurementTradeItemHeight">
          <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemHeightUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$BI/fnf_{1}:MeasurementTradeItemHeight"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$item/fnf_{1}:MeasurementTradeItemHeight"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1009</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$item/fnf_{1}:MeasurementTradeItemHeight">
          <mapTo list="conversion" tocode="MMT" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemHeightUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$item/fnf_{1}:MeasurementTradeItemHeight"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
    <Cell ss:StyleID="s291"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s295"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П25</Data></Cell>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21.75" ss:StyleID="s26">
    <Cell ss:StyleID="s290"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>Масса нетто                  </B><Font>гр.</Font></ss:Data></Cell>
    <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$BI/fnf_{1}:NetWeight"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1009</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$BI/fnf_{1}:NetWeight">
          <mapTo list="conversion" tocode="GRM" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$BI/fnf_{1}:NetWeightUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$BI/fnf_{1}:NetWeight"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$item/fnf_{1}:NetWeight"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1009</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$item/fnf_{1}:NetWeight">
          <mapTo list="conversion" tocode="GRM" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$item/fnf_{1}:NetWeightUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$item/fnf_{1}:NetWeight"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
    <Cell ss:StyleID="s291"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s295"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П26</Data></Cell>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="47.25" ss:StyleID="s26">
    <Cell ss:StyleID="s305"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B>Масса брутто                </B><Font>гр.</Font></ss:Data></Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$BI/fnf_{1}:GrossWeightFNFTSS"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1010</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$BI/fnf_{1}:GrossWeightFNFTSS">
          <mapTo list="conversion" tocode="GRM" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$BI/fnf_{1}:GrossWeightUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$BI/fnf_{1}:GrossWeightFNFTSS"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
     <Cell>
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$item/fnf_{1}:GrossWeightFNFTSS"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1010</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
      <Data ss:Type="Number">
        <xsl:if test="$item/fnf_{1}:GrossWeightFNFTSS">
          <mapTo list="conversion" tocode="GRM" firstcell="code" secondcell="tocode" thirdcell="coefficient">
            <xsl:value-of select="$item/fnf_{1}:GrossWeightUnitOfMeasure"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of select="$item/fnf_{1}:GrossWeightFNFTSS"/>
          </mapTo>
        </xsl:if>
      </Data>
    </Cell>
    <Cell ss:StyleID="s307"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s308"><Data ss:Type="String">не заполняется</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П27</Data></Cell>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21.75" ss:StyleID="s26">
    <Cell ss:StyleID="s309"/>
    <Cell ss:StyleID="s310"><Data ss:Type="String">ШТРИХ-КОД</Data></Cell>
    <Cell ss:StyleID="s310"/>
    <Cell ss:StyleID="s310"/>
    <Cell ss:StyleID="s310"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П28</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s312"><Data ss:Type="String">Аналог</Data></Cell>
    <Cell ss:StyleID="s316"><Data ss:Type="String">да</Data></Cell>
    <Cell ss:StyleID="s316"/>
    <Cell ss:StyleID="s317"><Data ss:Type="String">нет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ22</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="44.25" ss:StyleID="s29">
    <Cell ss:StyleID="s318"/>
    <Cell ss:StyleID="s319"><Data ss:Type="String">ЕИ</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20678274"><Data ss:Type="String">Штрих-код</Data></Cell>
    <Cell ss:Index="7" ss:MergeAcross="2" ss:StyleID="s328"><Data ss:Type="String">Для аналога: внутренний товарный код выводимой позиции:</Data></Cell>
    <Cell ss:StyleID="s332"/>
    <Cell ss:StyleID="s332"/>
    <Cell ss:StyleID="s333"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ23</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="41.25" ss:StyleID="s334">
    <Cell ss:StyleID="s337"><Data ss:Type="String">Штрих-код продукта:</Data></Cell>
    <Cell ss:StyleID="s338"><Data ss:Type="String">шт</Data></Cell>
    <Cell ss:MergeAcross="2">
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$BI/fnf_{1}:GTIN"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1011</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="Number">
       <xsl:value-of select="$BI/fnf_{1}:GTIN"/>
      </Data>
    </Cell>
    <Cell ss:StyleID="s346"/>
    <Cell ss:StyleID="s347"><Data ss:Type="String">Новинка</Data></Cell>
    <Cell ss:StyleID="s348"/>
    <Cell ss:StyleID="s348"/>
    <Cell ss:StyleID="s349"><Data ss:Type="String">да</Data></Cell>
    <Cell ss:StyleID="s349"/>
    <Cell ss:StyleID="s350"><Data ss:Type="String">нет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ24</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="20.25" ss:StyleID="s334">
    <Cell ss:StyleID="s351"><Data ss:Type="String">Штрих-код продукта:</Data></Cell>
    <Cell ss:StyleID="s338"><Data ss:Type="String">шт</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20670562"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П29</Data></Cell>
    <Cell ss:StyleID="s335"/>
    <Cell ss:Index="13" ss:StyleID="s336"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s29">
    <Cell ss:StyleID="s351"><Data ss:Type="String">Штрих-код продукта:</Data></Cell>
    <Cell ss:StyleID="s338"><Data ss:Type="String">Трансп.упак.</Data></Cell>
    <Cell ss:MergeAcross="2">
      <xsl:call-template name="markcellchanges">
        <xsl:with-param name="node" select="$item/fnf_{1}:GTIN"/>
        <xsl:with-param name="styleID">
          <xsl:text>s1012</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      <Data ss:Type="String">
        <xsl:value-of select="$item/fnf_{1}:GTIN"/>
      </Data>
     </Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П30</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s360"><Data ss:Type="String">Набор</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670704"><Data ss:Type="String">да</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670714"><Data ss:Type="String">нет</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ25</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="20.25" ss:StyleID="s29">
    <Cell ss:StyleID="s351"><Data ss:Type="String">Штрих-код продукта:</Data></Cell>
    <Cell ss:StyleID="s338"><Data ss:Type="String">Трансп.упак.</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20670724"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П31</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s380"/>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670744"><Data ss:Type="String">вн. товарный код</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670754"><Data ss:Type="String">Количество</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ26</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s29">
    <Cell ss:StyleID="s392"><Data ss:Type="String">Штрих-код продукта:</Data></Cell>
    <Cell ss:StyleID="s393"/>
    <Cell ss:StyleID="s306"/>
    <Cell ss:StyleID="s394"/>
    <Cell ss:StyleID="s395"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П32</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670856"><Data ss:Type="String">SKU1</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670866"/>
    <Cell ss:MergeAcross="1" ss:StyleID="m20670876"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ27</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s29">
    <Cell ss:MergeAcross="4" ss:StyleID="s46"><Data ss:Type="String">ВНУТРЕННЕЕ АРТИКУЛЬНОЕ ОБОЗНАЧЕНИЕ ПОСТАВЩИКА</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П33</Data></Cell>
    <Cell ss:StyleID="s410"><Data ss:Type="String">SKU2</Data></Cell>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s332"/>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s411"/>
    <Cell ss:StyleID="s333"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ28</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="31.5" ss:StyleID="s29">
    <Cell ss:MergeAcross="1" ss:StyleID="s414"><Data ss:Type="String" x:Ticked="1"> Внутренний номер товара у поставщика:</Data></Cell>
    <Cell ss:StyleID="s418"/>
     <Cell ss:MergeAcross="1">
       <xsl:call-template name="markcellchanges">
         <xsl:with-param name="node" select="$item/fnf_{1}:InternalItemIDofSupplier"/>
         <xsl:with-param name="styleID">
           <xsl:text>s1013</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
       <Data ss:Type="String">
         <xsl:value-of select="$item/fnf_{1}:InternalItemIDofSupplier"/>
       </Data>
     </Cell>
    <Cell ss:Index="7" ss:StyleID="s410"><Data ss:Type="String">SKU3</Data></Cell>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s332"/>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s411"/>
    <Cell ss:StyleID="s333"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ29</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s29">
    <Cell ss:MergeAcross="4" ss:StyleID="s424"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><I>ЗАКУПОЧНЫЕ ЦЕНЫ </I><B><I>БЕЗ</I></B><I> </I><B><I>НДС</I></B></ss:Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П34</Data></Cell>
    <Cell ss:StyleID="s410"><Data ss:Type="String">SKU4</Data></Cell>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s332"/>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s411"/>
    <Cell ss:StyleID="s333"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ30</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="20.25" ss:StyleID="s26">
    <Cell ss:StyleID="s426"><Data ss:Type="String">ЗЦ по прайс-листу:                  </Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20671038"/>
    <Cell ss:Index="7" ss:StyleID="s410"><Data ss:Type="String">SKU5</Data></Cell>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s332"/>
    <Cell ss:StyleID="s402"/>
    <Cell ss:StyleID="s411"/>
    <Cell ss:StyleID="s333"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ31</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:StyleID="s426"><Data ss:Type="String">Валюта:</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20671176"><Data ss:Type="String">рубли</Data></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П35</Data></Cell>
    <Cell ss:StyleID="s441"><Data ss:Type="String">SKU6</Data></Cell>
    <Cell ss:StyleID="s442"/>
    <Cell ss:StyleID="s443"/>
    <Cell ss:StyleID="s442"/>
    <Cell ss:StyleID="s444"/>
    <Cell ss:StyleID="s445"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">КМ32</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21.75" ss:StyleID="s29">
    <Cell ss:StyleID="s446"><Data ss:Type="String">Скидка:</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20671186"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П36</Data></Cell>
    <Cell ss:StyleID="s30"/>
    <Cell ss:Index="13" ss:StyleID="s31"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:StyleID="s454"><Data ss:Type="String">Цена для Ленты:</Data><Comment
      ss:Author="ekaterina.zarayskaya"><ss:Data
       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"
         x:CharSet="238" html:Size="8" html:Color="#000000">ekaterina.zarayskaya:</Font></B><Font
        html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#000000">&#10;заносится в SAP </Font></ss:Data></Comment></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20671196"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П37</Data></Cell>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:MergeAcross="4" ss:StyleID="s47"><Data ss:Type="String">СКИДКА</Data><Comment
      ss:Author="ekaterina.zarayskaya"><ss:Data
       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"
         x:CharSet="238" html:Size="8" html:Color="#000000">ekaterina.zarayskaya:</Font></B><Font
        html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#000000">&#10;</Font><B><Font
         html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#FF0000">главное - указать новую ЗЦ в руб и период ее действия для занесения в SAP </Font></B></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П38</Data></Cell>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="13" ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s457">
    <Cell ss:StyleID="s460"><Data ss:Type="String">Действует с:</Data></Cell>
    <Cell ss:StyleID="s461"/>
    <Cell ss:StyleID="s462"/>
    <Cell ss:StyleID="s462"/>
    <Cell ss:StyleID="s463"/>
    <Cell ss:Index="7" ss:StyleID="s458"/>
    <Cell ss:Index="13" ss:StyleID="s464"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21.75" ss:StyleID="s457">
    <Cell ss:StyleID="s465"><Data ss:Type="String">Действует по:</Data></Cell>
    <Cell ss:StyleID="s466"/>
    <Cell ss:StyleID="s467"/>
    <Cell ss:StyleID="s467"/>
    <Cell ss:StyleID="s468"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П39</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s470"><Data ss:Type="String">Оператор:</Data></Cell>
    <Cell ss:StyleID="s474"/>
    <Cell ss:StyleID="s475"/>
    <Cell ss:StyleID="s475"/>
    <Cell ss:StyleID="s476"/>
    <Cell ss:StyleID="s477"><Data ss:Type="String">О2</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="22.5" ss:StyleID="s457">
    <Cell ss:StyleID="s478"><Data ss:Type="String">ЗЦ без НДС со скидкой:</Data><Comment
      ss:Author="ekaterina.zarayskaya"><ss:Data
       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"
         x:CharSet="238" html:Size="8" html:Color="#000000">ekaterina.zarayskaya:</Font></B><Font
        html:Face="Tahoma" x:CharSet="238" html:Size="8" html:Color="#000000">&#10;заносится в SAP </Font></ss:Data></Comment></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20672338"/>
    <Cell ss:StyleID="s55"><Data ss:Type="String">П40</Data></Cell>
    <Cell ss:StyleID="s458"/>
    <Cell ss:Index="13" ss:StyleID="s464"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="20.25" ss:StyleID="s26">
    <Cell ss:StyleID="s484"/>
    <Cell ss:StyleID="s485"/>
    <Cell ss:StyleID="s485"/>
    <Cell ss:StyleID="s485"/>
    <Cell ss:StyleID="s55"/>
    <Cell ss:StyleID="s486"/>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="12" ss:StyleID="s464"/>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:StyleID="s26">
    <Cell ss:StyleID="s484"/>
    <Cell ss:StyleID="s485"/>
    <Cell ss:StyleID="s485"/>
    <Cell ss:StyleID="s485"/>
    <Cell ss:StyleID="s55"/>
    <Cell ss:StyleID="s486"/>
    <Cell ss:StyleID="s25"/>
    <Cell ss:Index="12" ss:StyleID="s464"/>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:StyleID="s487"><Data ss:Type="String">Заполняется в Ленте</Data></Cell>
    <Cell ss:StyleID="s488"/>
    <Cell ss:StyleID="s489"/>
    <Cell ss:StyleID="s488"/>
    <Cell ss:StyleID="s490"/>
    <Cell ss:StyleID="s489"/>
    <Cell ss:StyleID="s491"/>
    <Cell ss:StyleID="s491"/>
    <Cell ss:StyleID="s491"/>
    <Cell ss:StyleID="s491"/>
    <Cell ss:StyleID="s492"/>
    <Cell ss:StyleID="s493"/>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="30.75" ss:StyleID="s26">
    <Cell ss:StyleID="s494"><Data ss:Type="String">SKU действует в:</Data><Comment
      ss:Author="ekaterina.zarayskaya"><ss:Data
       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"
         x:CharSet="238" html:Size="12" html:Color="#000000">ТК, в которых данный товар должен быть в активной матрице</Font></B></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s495"><Data ss:Type="String">Л-1 Л-2 Л-3 Л-4 Л-5 Л-6 Л-7 Л-8 Л-9 Л-10 Л-11 Л-12 Л-14 Л-15 Л-16 Л-31 Л-32 Л-51 Л-52 Л-54 Л-71 Л-72 Л-73 Л-75 Л-76 Л-77 Л-91</Data></Cell>
    <Cell ss:StyleID="s496"/>
    <Cell ss:StyleID="s496"/>
    <Cell ss:StyleID="s496"/>
    <Cell ss:StyleID="s496"/>
    <Cell ss:StyleID="s497"/>
    <Cell ss:StyleID="s497"/>
    <Cell ss:StyleID="s497"/>
    <Cell ss:StyleID="s497"/>
    <Cell ss:StyleID="s498"/>
    <Cell ss:MergeDown="1" ss:StyleID="s477"><Data ss:Type="String">КМ29</Data></Cell>
    <Cell ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:MergeDown="1" ss:StyleID="m20672358"><Data ss:Type="String">Развозка через РЦ:</Data><Comment
      ss:Author="ekaterina.zarayskaya"><ss:Data
       xmlns="http://www.w3.org/TR/REC-html40"><Font html:Face="Tahoma"
        x:CharSet="238" html:Size="12" html:Color="#000000">Способ доставки товара в указанный регион</Font></ss:Data></Comment></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20672368"><Data ss:Type="String"
      x:Ticked="1">ЦРЦ - РРЦ - ТК</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m20672378"><Data ss:Type="String"
      x:Ticked="1">ЦРЦ - ТК</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20672388"><Data ss:Type="String">Напрямую</Data></Cell>
    <Cell ss:Index="13" ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21.75" ss:StyleID="s26">
    <Cell ss:Index="2" ss:StyleID="s530"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s531"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s530"/>
    <Cell ss:StyleID="s532"/>
    <Cell ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:StyleID="s533"/>
    <Cell ss:StyleID="s534"/>
    <Cell ss:StyleID="s534"/>
    <Cell ss:StyleID="s534"/>
    <Cell ss:StyleID="s534"/>
    <Cell ss:StyleID="s534"/>
    <Cell ss:StyleID="s534"/>
    <Cell ss:StyleID="s535"/>
    <Cell ss:StyleID="s536"/>
    <Cell ss:StyleID="s536"/>
    <Cell ss:StyleID="s537"/>
    <Cell ss:StyleID="s532"/>
    <Cell ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:MergeDown="1" ss:StyleID="m20673672"/>
    <Cell ss:MergeDown="1" ss:StyleID="m20673682"><Data ss:Type="String">ЦРЦ</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m20673692"><Data ss:Type="String">СПБ (ТК)</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20673702"><Data ss:Type="String">Новосибирск</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20673824"><Data ss:Type="String">Астрахань</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m20673834"><Data ss:Type="String">Тюмень</Data></Cell>
    <Cell ss:StyleID="s532"/>
    <Cell ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:Index="4" ss:StyleID="s565"><Data ss:Type="String">РРЦ</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20673844"><Data ss:Type="String">ТК</Data></Cell>
    <Cell ss:StyleID="s556"><Data ss:Type="String">РРЦ</Data></Cell>
    <Cell ss:StyleID="s556"><Data ss:Type="String">ТК</Data></Cell>
    <Cell ss:StyleID="s568"><Data ss:Type="String">РРЦ</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20673854"><Data ss:Type="String">ТК</Data></Cell>
    <Cell ss:Index="13" ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:StyleID="s575"><Data ss:Type="String">Дата первого заказа:</Data></Cell>
    <Cell ss:StyleID="s576"/>
    <Cell ss:StyleID="s577"/>
    <Cell ss:StyleID="s577"/>
    <Cell ss:MergeAcross="1" ss:StyleID="m20673976"/>
    <Cell ss:StyleID="s577"/>
    <Cell ss:StyleID="s556"/>
    <Cell ss:StyleID="s578"/>
    <Cell ss:MergeAcross="1" ss:StyleID="m20673986"/>
    <Cell ss:StyleID="s477"><Data ss:Type="String">КМ30</Data></Cell>
    <Cell ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:StyleID="s579"><Data ss:Type="String">Действует до:</Data></Cell>
    <Cell ss:StyleID="s576"/>
    <Cell ss:StyleID="s577"/>
    <Cell ss:StyleID="s577"/>
    <Cell ss:MergeAcross="1" ss:StyleID="m20673996"/>
    <Cell ss:StyleID="s577"/>
    <Cell ss:StyleID="s556"/>
    <Cell ss:StyleID="s578"/>
    <Cell ss:MergeAcross="1" ss:StyleID="m20674006"/>
    <Cell ss:StyleID="s477"><Data ss:Type="String">КМ31</Data></Cell>
    <Cell ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s26">
    <Cell ss:StyleID="s580"/>
    <Cell ss:StyleID="s581"/>
    <Cell ss:StyleID="s581"/>
    <Cell ss:StyleID="s581"/>
    <Cell ss:StyleID="s581"/>
    <Cell ss:StyleID="s581"/>
    <Cell ss:StyleID="s582"/>
    <Cell ss:StyleID="s582"/>
    <Cell ss:StyleID="s582"/>
    <Cell ss:StyleID="s582"/>
    <Cell ss:StyleID="s583"/>
    <Cell ss:StyleID="s584"/>
    <Cell ss:StyleID="s501"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="30" ss:StyleID="s26">
    <Cell ss:StyleID="s585"><Data ss:Type="String">Менеджер КС: .</Data></Cell>
    <Cell ss:MergeAcross="6" ss:StyleID="m20674128"/>
    <Cell ss:StyleID="s593"><Data ss:Type="String">Дата:</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="m20674138"/>
    <Cell ss:StyleID="s600"><Data ss:Type="String">КМ32</Data></Cell>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="36" ss:StyleID="s26">
    <Cell ss:MergeAcross="1" ss:StyleID="s602"><Data ss:Type="String">Решение Директора направления: </Data></Cell>
    <Cell ss:StyleID="s606"/>
    <Cell ss:StyleID="s606"/>
    <Cell ss:StyleID="s606"/>
    <Cell ss:StyleID="s606"/>
    <Cell ss:StyleID="s606"/>
    <Cell ss:StyleID="s606"/>
    <Cell ss:StyleID="s607"/>
    <Cell ss:StyleID="s607"/>
    <Cell ss:StyleID="s608"/>
    <Cell ss:StyleID="s609"/>
    <Cell ss:StyleID="s27"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="31.5">
    <Cell ss:Index="4" ss:StyleID="s21"/>
    <Cell ss:Index="13" ss:StyleID="s55"><Data ss:Type="String">КМ40</Data></Cell>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
   <Row>
    <Cell ss:Index="4" ss:StyleID="s21"/>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.31496062992125984"/>
    <Footer x:Margin="0.31496062992125984"/>
    <PageMargins x:Bottom="0.19685039370078741" x:Left="0.59055118110236227"
     x:Right="0.19685039370078741" x:Top="0.19685039370078741"/>
   </PageSetup>
   <FitToPage/>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <Scale>38</Scale>
    <HorizontalResolution>300</HorizontalResolution>
    <VerticalResolution>300</VerticalResolution>
   </Print>
   <Zoom>55</Zoom>
   <PageBreakZoom>50</PageBreakZoom>
   <Selected/>
   <DoNotDisplayGridlines/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>18</ActiveRow>
     <ActiveCol>1</ActiveCol>
     <RangeSelection>R19C2:R19C4</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
   <Range>R12C10:R12C11</Range>
   <Type>TextLength</Type>
   <Qualifier>LessOrEqual</Qualifier>
   <Value>20</Value>
   <InputTitle>Povolená délka názvu 20 znaků</InputTitle>
   <InputMessage>Krátký název výrobku může obsahovat pouze 20 znaků vč. interpunkcí a mezer</InputMessage>
   <ErrorMessage>Zadali jste</ErrorMessage>
   <ErrorTitle>Chybná délka názvu</ErrorTitle>
  </DataValidation>
 </Worksheet>
</Workbook>
  </xsl:template>
  <xsl:template match="/">
    <xsl:for-each select="fnf_{1}:Item">
      <xsl:variable name="BI" select="fnf_{1}:BaseItem/fnf_{1}:BaseItemVersion[fnf_{1}:GTIN='{3}']" />
      <xsl:variable name="BA" select="fnf_{1}:Assortment/fnf_{1}:AssortmentVersion[fnf_{1}:GTIN='{3}']" />
      <xsl:if test="$BI">
        <xsl:call-template name="dataout">
          <xsl:with-param name="item" select="fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTIN='{4}']"/>
          <xsl:with-param name="BI" select="$BI"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test="$BA">
        <xsl:call-template name="dataout">
          <xsl:with-param name="item" select="fnf_{1}:PackagingItem/fnf_{1}:PackagingItemVersion[fnf_{1}:GTIN='{4}']"/>
          <xsl:with-param name="BI" select="$BA"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:for-each>
   </xsl:template>
</xsl:stylesheet>