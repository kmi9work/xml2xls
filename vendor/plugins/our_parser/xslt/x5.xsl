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
        <Author>Baglukov</Author>
        <LastAuthor>Rogozhin Anton Yurevich</LastAuthor>
        <LastPrinted>2009-07-16T09:15:16Z</LastPrinted>
        <Created>2001-08-16T06:13:51Z</Created>
        <LastSaved>2009-10-17T22:56:15Z</LastSaved>
        <Company>Perekriostok</Company>
        <Version>12.00</Version>
      </DocumentProperties>
      <CustomDocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
        <_NewReviewCycle dt:dt="string"></_NewReviewCycle>
      </CustomDocumentProperties>
      <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
        <WindowHeight>9015</WindowHeight>
        <WindowWidth>15165</WindowWidth>
        <WindowTopX>0</WindowTopX>
        <WindowTopY>1365</WindowTopY>
        <TabRatio>426</TabRatio>
        <ActiveSheet>1</ActiveSheet>
        <FirstVisibleSheet>1</FirstVisibleSheet>
        <ProtectStructure>False</ProtectStructure>
        <ProtectWindows>False</ProtectWindows>
      </ExcelWorkbook>
      <Styles>
        <Style s:ID="Default" s:Name="Normal">
          <Alignment s:Vertical="Bottom"/>
          <Borders/>
          <Font s:FontName="Arial" x:CharSet="204"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style s:ID="s16" s:Name="0,0&#13;&#10;NA&#13;&#10;">
          <Alignment s:Vertical="Bottom"/>
          <Borders/>
          <Font s:FontName="Arial" x:CharSet="204"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style s:ID="s35" s:Name="Comma [0]_Actual Volume_00">
          <NumberFormat s:Format="_-* #,##0_$_-;\-* #,##0_$_-;_-* &quot;-&quot;_$_-;_-@_-"/>
        </Style>
        <Style s:ID="s36" s:Name="Normal_assortment">
          <Alignment s:Vertical="Bottom"/>
          <Borders/>
          <Font s:FontName="Arial" x:CharSet="204"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style s:ID="s37" s:Name="Normal_coffee local price list1">
          <Alignment s:Vertical="Bottom"/>
          <Borders/>
          <Font s:FontName="Arial" x:CharSet="204"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style s:ID="s38" s:Name="Normal_Sheet1">
          <Alignment s:Vertical="Bottom"/>
          <Borders/>
          <Font s:FontName="Arial"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style s:ID="s59" s:Name="Обычный_Лист1">
          <Alignment s:Horizontal="Left" s:Vertical="Bottom"/>
          <Borders/>
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Size="8"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style s:ID="m38548352" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548372" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548412" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548432" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548452" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548472" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548492" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548128" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548148" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548168" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548188" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548208" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38548228" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38548268" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38547924" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38547944" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38547964" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38547984" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="m38548004" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38547700" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38547840" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38547476" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38547496" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38547596" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38547616" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051552" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051592" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051612" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051632" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051652" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051672" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051692" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38051328" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051348" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051368" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38051388" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="m38051408" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051428" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051448" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051468" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="m38051508" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s70" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s71" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s72" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s73" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s74" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s75" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s76" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s77" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s78" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s79" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s80" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s81" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s82" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s83" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s84" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s85" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s86" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s87" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s88" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s89" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s90" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top"/>
          <Borders/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s91" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s92" s:Parent="s16">
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"/>
          <Interior/>
        </Style>
        <Style s:ID="s93" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s94" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s95" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s96" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s97" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Borders/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s98" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#000000" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s99" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat s:Format="Fixed"/>
        </Style>
        <Style s:ID="s100" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s101" s:Parent="s59">
          <Font s:FontName="Arial" x:Family="Swiss"/>
        </Style>
        <Style s:ID="s102" s:Parent="s59">
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
        </Style>
        <Style s:ID="s103" s:Parent="s16">
          <Font s:FontName="Arial" x:Family="Swiss"/>
        </Style>
        <Style s:ID="s104" s:Parent="s59">
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Bold="1"/>
        </Style>
        <Style s:ID="s105" s:Parent="s16">
          <Font s:FontName="Arial"/>
        </Style>
        <Style s:ID="s106" s:Parent="s16">
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="12" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s107" s:Parent="s16">
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="11" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s108" s:Parent="s16">
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="11" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s109" s:Parent="s16">
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="11" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s110" s:Parent="s16">
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="11" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s111" s:Parent="s16">
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="11" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s112" s:Parent="s16">
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="11" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s113" s:Parent="s16">
          <Font s:FontName="Georgia" x:Family="Roman" s:Size="11" s:Color="#FF0000"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s114" s:Parent="s16">
          <Font s:FontName="Times New Roman" x:Family="Roman" s:Size="12"/>
        </Style>
        <Style s:ID="s115" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s116" s:Parent="s16">
          <Alignment s:Horizontal="Right" s:Vertical="Top" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s117" s:Parent="s16">
          <Alignment s:Vertical="Center"/>
          <Borders/>
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Size="14"
           s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s118" s:Parent="s16">
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Size="12"
           s:Color="#FF0000" s:Bold="1" s:Italic="1"/>
        </Style>
        <Style s:ID="s119" s:Parent="s16">
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Size="12"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s120" s:Parent="s16">
          <Font s:FontName="Arial" s:Size="12"/>
        </Style>
        <Style s:ID="s121" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12"/>
          <Interior/>
        </Style>
        <Style s:ID="s122" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12"/>
          <Interior/>
          <NumberFormat s:Format="Fixed"/>
        </Style>
        <Style s:ID="s123" s:Parent="s16">
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="12"/>
          <Interior/>
        </Style>
        <Style s:ID="s124" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s125" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12"/>
          <Interior/>
        </Style>
        <Style s:ID="s126" s:Parent="s16">
          <Alignment s:Horizontal="Right" s:Vertical="Top" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12"/>
          <Interior/>
        </Style>
        <Style s:ID="s127" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12" s:Color="#FF0000"
           s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s128" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12" s:Color="#FF0000"
           s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s129" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12" s:Color="#FF0000"
           s:Bold="1"/>
          <Interior/>
          <NumberFormat s:Format="Fixed"/>
        </Style>
        <Style s:ID="s130" s:Parent="s16">
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="12" s:Color="#FF0000" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s131" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s132" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s133" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s134" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Borders/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s135" s:Parent="s59">
          <Alignment s:Horizontal="Left" s:Vertical="Bottom"/>
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Bold="1"/>
        </Style>
        <Style s:ID="s136" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Bottom"/>
          <Font s:FontName="Arial" x:CharSet="204" x:Family="Swiss" s:Size="12"
           s:Bold="1"/>
        </Style>
        <Style s:ID="s137" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12"/>
          <Interior/>
        </Style>
        <Style s:ID="s138" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s139" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s140" s:Parent="s59">
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="22" s:Bold="1"/>
        </Style>
        <Style s:ID="s141" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s142" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s143" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s144" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s145" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s146" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s147" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s148" s:Parent="s16">
          <Alignment s:Horizontal="Left" s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s149" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s150" s:Parent="s16">
          <Alignment s:Vertical="Top" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s151" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s152" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s153" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s154" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s155" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s156" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s157" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s158" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s159" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:Rotate="90"
           s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior s:Color="#FFFF00" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s160" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12"/>
          <Interior/>
        </Style>
        <Style s:ID="s161" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Font s:FontName="Arial" x:Family="Swiss" s:Size="12"/>
          <Interior/>
        </Style>
        <Style s:ID="s162">
          <Alignment s:Horizontal="Center" s:Vertical="Bottom" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
          </Borders>
          <Font s:FontName="Arial Cyr" x:CharSet="204" s:Size="8" s:Color="#000000"/>
          <Interior/>
          <NumberFormat s:Format="0000;&quot;&quot;;&quot;&quot;"/>
        </Style>
        <Style s:ID="s163" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s164" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s165" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss" s:Bold="1"/>
          <Interior/>
        </Style>
        <Style s:ID="s166" s:Parent="s35">
          <Alignment s:Horizontal="Left" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s167" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
        </Style>
        <Style s:ID="s168">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14" s:Color="#000000"/>
          <Interior/>
          <NumberFormat s:Format="0000;&quot;&quot;;&quot;&quot;"/>
        </Style>
        <Style s:ID="s169" s:Parent="s38">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
        </Style>
        <Style s:ID="s170" s:Parent="s36">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="@"/>
        </Style>
        <Style s:ID="s171" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s172">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="0.000"/>
        </Style>
        <Style s:ID="s173">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s174">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
        </Style>
        <Style s:ID="s175" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior />
        </Style>
        <Style s:ID="s176" s:Parent="s36">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s177">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior/>
        </Style>
        <Style s:ID="s178" s:Parent="s16">
          <Alignment s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior/>
        </Style>
        <Style s:ID="s179" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior/>
        </Style>
        <Style s:ID="s180" s:Parent="s38">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior s:Color="#FFFFFF" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s181" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s182">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior/>
        </Style>
        <Style s:ID="s183" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior s:Color="#FFFFFF" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s184" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior/>
        </Style>
        <Style s:ID="s185" s:Parent="s35">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s186" s:Parent="s35">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="15"/>
          <Interior/>
          <NumberFormat s:Format="@"/>
        </Style>
        <Style s:ID="s187" s:Parent="s37">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior s:Color="#FF99CC" s:Pattern="Solid"/>
          <NumberFormat s:Format="Standard"/>
        </Style>
        <Style s:ID="s188" s:Parent="s37">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior s:Color="#FF99CC" s:Pattern="Solid"/>
          <NumberFormat s:Format="Standard"/>
        </Style>
        <Style s:ID="s228" s:Parent="s16">
          <Alignment s:Vertical="Top"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="2"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s270" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Top" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
        <Style s:ID="s271" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior s:Color="#888888" s:Pattern="Solid"/>
        </Style>
        <Style s:ID="s1001" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior/>
        </Style>
       <Style s:ID="s1002" s:Parent="s35">
          <Alignment s:Horizontal="Left" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat/>
        </Style>
      <Style s:ID="s1003" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
        </Style>
        <Style s:ID="s1004">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14" s:Color="#000000"/>
          <Interior/>
          <NumberFormat s:Format="0000;&quot;&quot;;&quot;&quot;"/>
        </Style>
        <Style s:ID="s1005" s:Parent="s38">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
        </Style>
        <Style s:ID="s1006" s:Parent="s36">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="@"/>
        </Style>
        <Style s:ID="s1007" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s1008">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="0.000"/>
        </Style>
        <Style s:ID="s1009">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s1010">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
        </Style>
        <Style s:ID="s1011" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior />
        </Style>
       <Style s:ID="s1012" s:Parent="s35">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior/>
          <NumberFormat/>
        </Style>
       <Style s:ID="s1013" s:Parent="s37">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior s:Color="#FF99CC" s:Pattern="Solid"/>
          <NumberFormat s:Format="Standard"/>
        </Style>
        <Style s:ID="s2001" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s2002" s:Parent="s35">
          <Alignment s:Horizontal="Left" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s2003" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s2004">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14" s:Color="#000000"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0000;&quot;&quot;;&quot;&quot;"/>
        </Style>
        <Style s:ID="s2005" s:Parent="s38">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s2006" s:Parent="s36">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="@"/>
        </Style>
        <Style s:ID="s2007" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s2008">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0.000"/>
        </Style>
        <Style s:ID="s2009">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s2010">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
        </Style>
        <Style s:ID="s2011" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s2012" s:Parent="s35">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s2013" s:Parent="s37">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="Standard"/>
        </Style>
        <Style s:ID="s3001" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s3002" s:Parent="s35">
          <Alignment s:Horizontal="Left" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s3003" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s3004">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14" s:Color="#000000"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0000;&quot;&quot;;&quot;&quot;"/>
        </Style>
        <Style s:ID="s3005" s:Parent="s38">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s3006" s:Parent="s36">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="@"/>
        </Style>
        <Style s:ID="s3007" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s3008">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0.000"/>
        </Style>
        <Style s:ID="s3009">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s3010">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
        </Style>
        <Style s:ID="s3011" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s3012" s:Parent="s35">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s3013" s:Parent="s37">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat s:Format="Standard"/>
        </Style>
        <Style s:ID="s4001" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Center" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Arial" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s4002" s:Parent="s35">
          <Alignment s:Horizontal="Left" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s4003" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s4004">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"
             s:Color="#000000"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14" s:Color="#000000"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0000;&quot;&quot;;&quot;&quot;"/>
        </Style>
        <Style s:ID="s4005" s:Parent="s38">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s4006" s:Parent="s36">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat s:Format="@"/>
        </Style>
        <Style s:ID="s4007" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s4008">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0.000"/>
        </Style>
        <Style s:ID="s4009">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat s:Format="0"/>
        </Style>
        <Style s:ID="s4010">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
        </Style>
        <Style s:ID="s4011" s:Parent="s16">
          <Alignment s:Horizontal="Center" s:Vertical="Justify" s:WrapText="1"/>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
        </Style>
        <Style s:ID="s4012" s:Parent="s35">
          <Alignment s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat/>
        </Style>
        <Style s:ID="s4013" s:Parent="s37">
          <Alignment s:Horizontal="Center" s:Vertical="Justify"/>
          <Borders>
            <Border s:Position="Bottom" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Left" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Right" s:LineStyle="Continuous" s:Weight="1"/>
            <Border s:Position="Top" s:LineStyle="Continuous" s:Weight="1"/>
          </Borders>
          <Font s:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
           s:Size="14"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat s:Format="Standard"/>
        </Style>
      </Styles>
      <Names>
        <NamedRange s:Name="орпор" s:RefersTo="=сокращения!R1:R65536"/>
      </Names>
      <Worksheet s:Name="сокращения">
        <Table  s:StyleID="s105">
          <Column s:StyleID="s105" s:AutoFitWidth="0" s:Width="234" s:Span="2"/>
          <Row s:Height="15.75">
            <Cell s:StyleID="s106">
              <s:Data s:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40">
                <B>
                  <Font html:Color="#FF0000">б/к – без кости (косточек)</Font>
                </B>
                <Font
       html:Face="Times New Roman" x:Family="Roman"> </Font>
              </s:Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s107">
              <Data s:Type="String">г/к – горячего копчения</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s108">
              <Data s:Type="String">с/г – с головой</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s109">
              <Data s:Type="String">с/к – с костью (косточкой)</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">х/к – холодного копчения</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">б/г – без головы</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s111">
              <Data s:Type="String">п/ф – полуфабрикат</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">сл/с – слабосоленая</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">с/м – свежезамороженный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s108">
              <Data s:Type="String">п/уп – подарочная упаковка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s111">
              <Data s:Type="String">н/к – на коже</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">б/ш – без шкуры</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">кр. – красное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s107">
              <Data s:Type="String">шок.конф. – шоколадные конфеты</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">н/ш – на шкуре</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">сух. – сухое</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">наб.конф. – набор конфет</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">б/з – быстрозамороженный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">мар. – марочное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s111">
              <Data s:Type="String">жев.рез. – жевательная резинка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">пл/ст – пластиковый стакан</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">п/дес. – полудесертное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s107">
              <Data s:Type="String">твор.- твороженный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">в/м – варено-мороженный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">п/сух. – полусухое</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">стер. – стерелизованное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">вар.- вареная</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">п/сл. – полусладкое</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">т/п – тетра-пак</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">с/к – сырокопченая</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">мус. – мускатное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">гл. – глазированный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">в/к – варено-копченая</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">бел. – белое</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">негл. – неглазированный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">н/о – натуральная оболочка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">игр. – игристое</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">пл. – плавленый</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">б/о – белковая оболочка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">роз. – розовое</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">т/р – тетра-рекс</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">к/з – копчено-запеченная</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">к/уп – картонная упаковка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s109">
              <Data s:Type="String">к/м – кисломолочный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">п/к – полукопченная</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">фл. – фляжка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">тв. – твердый</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s112">
              <Data s:Type="String">в обс. – в обсыпке</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">кер. – керамика</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">терм. -&#160; термизированный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s108">
              <Data s:Type="String">ж/б – железная банка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">мет.кор. – металлическая коробка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s112">
              <Data s:Type="String">п/ст. – пластиковый стакан</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">ст/б – стеклянная банка (бутылка)</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s112">
              <Data s:Type="String">кувш. – кувшин</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s108">
              <Data s:Type="String">КПБ – комплект постельного белья</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">в/с – высший сорт</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s108">
              <Data s:Type="String">Св. – светлое</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">раст. – растение</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">в/уп – вакуумная упаковка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">тем. – темное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">ср. – средство</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">раст. – растворимый</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">пл/б – пластиковая бутылка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">ст. – стиральный</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">б/р – быстрорастворимый</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">бут. – бутылочное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">пмм – посудомоечная машина</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">б/пр – быстрого приготовления</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s110">
              <Data s:Type="String">мин. – минеральная</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">наб. – набор</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">в с/с – в собственном соку</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:StyleID="s112">
              <Data s:Type="String">б/алк. – безалкогольное</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">игр. – игрушка</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s110">
              <Data s:Type="String">пак. – пакет</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:Index="2" s:StyleID="s112">
              <Data s:Type="String">осв.возд. – освежитель воздуха</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
            <Cell s:StyleID="s112">
              <Data s:Type="String">в т/с – в томатном соусе</Data>
              <NamedCell
      s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Height="14.25">
            <Cell s:Index="2" s:StyleID="s113">
              <NamedCell s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Index="30" s:Height="14.25">
            <Cell s:StyleID="s113">
              <NamedCell s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Index="47" s:Height="14.25">
            <Cell s:StyleID="s113">
              <NamedCell s:Name="орпор"/>
            </Cell>
          </Row>
          <Row s:Index="84" s:Height="15.75">
            <Cell s:StyleID="s114">
              <NamedCell s:Name="орпор"/>
            </Cell>
          </Row>
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
          <PageSetup>
            <Layout x:Orientation="Landscape"/>
          </PageSetup>
          <Print>
            <ValidPrinterInfo/>
            <PaperSizeIndex>9</PaperSizeIndex>
            <Scale>98</Scale>
            <HorizontalResolution>600</HorizontalResolution>
            <VerticalResolution>600</VerticalResolution>
          </Print>
          <PageBreakZoom>60</PageBreakZoom>
          <Panes>
            <Pane>
              <Number>3</Number>
              <ActiveRow>15</ActiveRow>
            </Pane>
          </Panes>
          <ProtectObjects>False</ProtectObjects>
          <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
      </Worksheet>
      <Worksheet s:Name="Форма о товаре">
        <Names>
          <NamedRange s:Name="_FilterDatabase"
           s:RefersTo="='Форма о товаре'!R15C1:R26C38" s:Hidden="1"/>
        </Names>
        <Table  s:StyleID="s71">
          <Column s:StyleID="s115" s:AutoFitWidth="0" s:Width="16.5"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="72.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="24"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="116.25"/>
          <Column s:StyleID="s94" s:AutoFitWidth="0" s:Width="353.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="28.5"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="45"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="47.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="27.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="33.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="31.5"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="24.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="123.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="30.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="49.5"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="26.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="42.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="62.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="41.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="56.25"/>
          <Column s:StyleID="s71" s:Width="44.25" s:Span="1"/>
          <Column s:Index="23" s:StyleID="s71" s:AutoFitWidth="0" s:Width="27.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="33" s:Span="1"/>
          <Column s:Index="26" s:StyleID="s71" s:AutoFitWidth="0" s:Width="41.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="69"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="40.5"/>
          <Column s:StyleID="s71" s:Hidden="1" s:AutoFitWidth="0" s:Width="21.75"/>
          <Column s:StyleID="s71" s:Width="89.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="59.25"/>
          <Column s:StyleID="s74" s:AutoFitWidth="0" s:Width="168.75"/>
          <Column s:StyleID="s74" s:AutoFitWidth="0" s:Width="33"/>
          <Column s:StyleID="s74" s:AutoFitWidth="0" s:Width="32.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="63.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="30.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="43.5"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="31.5"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="39.75"/>
          <Column s:Index="41" s:StyleID="s71" s:AutoFitWidth="0" s:Width="21.75"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="26.25"/>
          <Column s:StyleID="s71" s:AutoFitWidth="0" s:Width="28.5" s:Span="1"/>
          <Column s:Index="45" s:StyleID="s71" s:AutoFitWidth="0" s:Width="26.25"/>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:MergeAcross="34" s:StyleID="s70">
              <Data s:Type="String">Форма предоставления информации о товаре поставщиками Компании Х5</Data>
            </Cell>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:Index="2" s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s91"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s70"/>
            <Cell s:StyleID="s72"/>
            <Cell s:StyleID="s72"/>
            <Cell s:StyleID="s72"/>
            <Cell s:StyleID="s90"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="39" s:StyleID="s70"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s141">
              <Data s:Type="String">Поставщик</Data>
            </Cell>
            <Cell s:StyleID="s142"/>
            <Cell s:StyleID="s142"/>
            <Cell s:StyleID="s142"/>
            <Cell s:StyleID="s143"/>
            <Cell s:StyleID="s142"/>
            <Cell s:StyleID="s142"/>
            <Cell s:Index="9">
              <Data s:Type="String">Курс пересчета</Data>
            </Cell>
            <Cell s:Index="19" s:StyleID="s73">
              <Data s:Type="String">Объекты торговой сети для прайс-листа (множество)</Data>
            </Cell>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="35" s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s144"/>
            <Cell s:StyleID="s145">
              <Data s:Type="String">N в БД Перекресток</Data>
            </Cell>
            <Cell s:StyleID="s145"/>
            <Cell s:StyleID="s145"/>
            <Cell s:StyleID="s146"/>
            <Cell s:StyleID="s145"/>
            <Cell s:StyleID="s147"/>
            <Cell s:Index="9" s:StyleID="s75">
              <Data s:Type="String"> </Data>
            </Cell>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s77"/>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38051692"/>
            <Cell s:StyleID="s98"/>
            <Cell s:MergeAcross="6" s:StyleID="m38051408">
              <Data s:Type="String">МОСКВА</Data>
            </Cell>
            <Cell s:StyleID="s78"/>
            <Cell s:StyleID="s79"/>
            <Cell s:StyleID="s80"/>
            <Cell s:StyleID="s100"/>
            <Cell s:StyleID="s117"/>
            <Cell s:StyleID="s117"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s148"/>
            <Cell s:StyleID="s149">
              <Data s:Type="String">N в БД Пятерочка</Data>
            </Cell>
            <Cell s:StyleID="s149"/>
            <Cell s:StyleID="s149"/>
            <Cell s:StyleID="s150"/>
            <Cell s:StyleID="s149"/>
            <Cell s:StyleID="s151"/>
            <Cell s:Index="9" s:StyleID="s81"/>
            <Cell s:StyleID="s82"/>
            <Cell s:StyleID="s82"/>
            <Cell s:StyleID="s82"/>
            <Cell s:StyleID="s82"/>
            <Cell s:StyleID="s83"/>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38051388"/>
            <Cell s:StyleID="s84"/>
            <Cell s:MergeAcross="6" s:StyleID="m38051428">
              <Data s:Type="String">РЯЗАНЬ</Data>
            </Cell>
            <Cell s:StyleID="s84"/>
            <Cell s:StyleID="s85"/>
            <Cell s:StyleID="s86"/>
            <Cell s:StyleID="s100"/>
            <Cell s:StyleID="s117"/>
            <Cell s:StyleID="s117"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell>
              <Data s:Type="String">Период согласования прайс-листа (ассортимент, цена)</Data>
            </Cell>
            <Cell s:Index="7" s:StyleID="s73"/>
            <Cell s:Index="9">
              <Data s:Type="String">Место поставки</Data>
            </Cell>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38548004"/>
            <Cell s:StyleID="s84"/>
            <Cell s:MergeAcross="6" s:StyleID="m38547924">
              <Data s:Type="String">ВЛАДИМИР</Data>
            </Cell>
            <Cell s:StyleID="s84"/>
            <Cell s:StyleID="s85"/>
            <Cell s:StyleID="s86"/>
            <Cell s:StyleID="s100"/>
            <Cell s:StyleID="s117"/>
            <Cell s:StyleID="s117"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s132"/>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s95"/>
            <Cell s:StyleID="s77">
              <Data s:Type="String"> </Data>
            </Cell>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="9" s:MergeAcross="1" s:MergeDown="1" s:StyleID="m38051592">
              <Data
      s:Type="String">магазин</Data>
            </Cell>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s76"/>
            <Cell s:StyleID="s76"/>
            <Cell s:MergeDown="1" s:StyleID="m38051612"/>
            <Cell s:MergeDown="1" s:StyleID="m38051632">
              <Data s:Type="String">РЦ</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38051652"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38051508"/>
            <Cell s:StyleID="s84"/>
            <Cell s:MergeAcross="6" s:StyleID="m38051468">
              <Data s:Type="String">ТУЛА</Data>
            </Cell>
            <Cell s:StyleID="s84"/>
            <Cell s:StyleID="s85"/>
            <Cell s:StyleID="s86"/>
            <Cell s:StyleID="s100"/>
            <Cell s:StyleID="s117"/>
            <Cell s:StyleID="s117"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s133">
              <Data s:Type="String"> </Data>
            </Cell>
            <Cell s:StyleID="s82">
              <Data s:Type="String"> </Data>
            </Cell>
            <Cell s:StyleID="s82"/>
            <Cell s:StyleID="s82"/>
            <Cell s:StyleID="s96"/>
            <Cell s:StyleID="s83"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="11" s:StyleID="s82"/>
            <Cell s:StyleID="s82"/>
            <Cell s:StyleID="s82"/>
            <Cell s:Index="17" s:StyleID="s73"/>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38051672"/>
            <Cell s:StyleID="s84"/>
            <Cell s:MergeAcross="6" s:StyleID="m38051448">
              <Data s:Type="String">ЯРОСЛАВЛЬ</Data>
            </Cell>
            <Cell s:StyleID="s84"/>
            <Cell s:StyleID="s85"/>
            <Cell s:StyleID="s86"/>
            <Cell s:StyleID="s100"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s134"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s97"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="9" s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38051552"/>
            <Cell s:StyleID="s163"/>
            <Cell s:MergeAcross="6" s:StyleID="s228">
              <Data s:Type="String">КАЛУГА</Data>
            </Cell>
            <Cell s:StyleID="s163"/>
            <Cell s:StyleID="s164"/>
            <Cell s:StyleID="s165"/>
            <Cell s:StyleID="s100"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s134"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s97"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="9" s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38051328"/>
            <Cell s:StyleID="s163"/>
            <Cell s:MergeAcross="6" s:StyleID="s228">
              <Data s:Type="String">ТВЕРЬ</Data>
            </Cell>
            <Cell s:StyleID="s163"/>
            <Cell s:StyleID="s164"/>
            <Cell s:StyleID="s165"/>
            <Cell s:StyleID="s100"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:StyleID="s134"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s97"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="9" s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:Index="19" s:MergeAcross="3" s:StyleID="m38547944"/>
            <Cell s:StyleID="s87"/>
            <Cell s:MergeAcross="6" s:StyleID="m38051348">
              <Data s:Type="String">ХРАМ ТОРГОВЛИ</Data>
            </Cell>
            <Cell s:StyleID="s87"/>
            <Cell s:StyleID="s88"/>
            <Cell s:StyleID="s89"/>
            <Cell s:StyleID="s100"/>
            <Cell s:Index="39" s:StyleID="s73"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="13.5">
            <Cell s:Index="22" s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s90"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
            <Cell s:StyleID="s73"/>
          </Row>
          <Row s:StyleID="s157">
            <Cell s:StyleID="s154">
              <Data s:Type="Number">1</Data>
            </Cell>
            <Cell s:StyleID="s154" s:Formula="=RC[-1]+1">
              <Data s:Type="Number">2</Data>
            </Cell>
            <Cell s:StyleID="s154" s:Formula="=RC[-1]+1">
              <Data s:Type="Number">3</Data>
            </Cell>
            <Cell s:StyleID="s154" s:Formula="=RC[-1]+1">
              <Data s:Type="Number">4</Data>
            </Cell>
            <Cell s:StyleID="s154" s:Formula="=RC[-1]+1">
              <Data s:Type="Number">5</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">6</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">7</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">8</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">9</Data>
            </Cell>
            <Cell s:StyleID="s156">
              <Data s:Type="Number">10</Data>
            </Cell>
            <Cell s:StyleID="s152">
              <Data s:Type="Number">11</Data>
            </Cell>
            <Cell s:StyleID="s153">
              <Data s:Type="Number">12</Data>
            </Cell>
            <Cell s:StyleID="s153">
              <Data s:Type="Number">13</Data>
            </Cell>
            <Cell s:StyleID="s153">
              <Data s:Type="Number">14</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">15</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">16</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">17</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">18</Data>
            </Cell>
            <Cell s:MergeAcross="3" s:StyleID="s154">
              <Data s:Type="Number">19</Data>
            </Cell>
            <Cell s:MergeAcross="2" s:StyleID="s154">
              <Data s:Type="Number">20</Data>
            </Cell>
            <Cell s:StyleID="s152">
              <Data s:Type="Number">21</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">22</Data>
            </Cell>
            <Cell s:MergeAcross="2" s:StyleID="m38547700">
              <Data s:Type="Number">23</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">24</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">25</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">26</Data>
            </Cell>
            <Cell s:StyleID="s152">
              <Data s:Type="Number">27</Data>
            </Cell>
            <Cell s:StyleID="s152">
              <Data s:Type="Number">28</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="s154">
              <Data s:Type="Number">29</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">30</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="m38548228">
              <Data s:Type="Number">31</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">32</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">33</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="Number">34</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="m38548352">
              <Data s:Type="Number">35</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="m38548128">
              <Data s:Type="Number">36</Data>
            </Cell>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="52.5" s:StyleID="s157">
            <Cell s:MergeDown="1" s:StyleID="s158">
              <Data s:Type="String">Порядковый номер</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s158">
              <Data s:Type="String">PLU Пятерочка</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s158">
              <Data s:Type="String">PLU Перекресток</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s138">
              <Data s:Type="String">Торговая марка на языке оригинала (ТМ)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s138">
              <Data s:Type="String">Наименование товара, свойства и характеристики</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38547840">
              <Data s:Type="String">Фин код</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s155">
              <Data s:Type="String">Вес, объем, емкость</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s155">
              <Data s:Type="String">Единица измерения          (шт, кг)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38547964">
              <Data s:Type="String">%содержания этил.спирта</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38547984">
              <Data s:Type="String">Код ЕГАИС</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38548168">
              <Data s:Type="String">Фасованный</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38548188">
              <Data s:Type="String">Маркированный</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38547596">
              <Data s:Type="String">штриховой код (коды) товара</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38547616">
              <Data s:Type="String">Внешний код поставщика (артикул поставщика. Производителя)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38547476">
              <Data s:Type="String">Цена единицы товара (плюс НДС)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s155">
              <Data s:Type="String">Ценовой сегмент &#10;(низкий, средний, высокий)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38547496">
              <Data s:Type="String">Количество единиц (шт/кг) в упаковке</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s155">
              <Data s:Type="String">Минимальный квант поставки (шт,кг)</Data>
            </Cell>
            <Cell s:MergeAcross="3" s:StyleID="s154">
              <Data s:Type="String">Параметры укладки на EURO-паллете</Data>
            </Cell>
            <Cell s:MergeAcross="2" s:StyleID="s154">
              <Data s:Type="String">Габариты товара, мм</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38548412">
              <Data s:Type="String">Кол-во упак.на паллете</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s138">
              <Data s:Type="String">Производитель товара (на языке оригинала)</Data>
            </Cell>
            <Cell s:MergeAcross="2" s:StyleID="m38548208">
              <Data s:Type="String">Группа складской аналитики</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s138">
              <Data s:Type="String">Код ОКДП/ТНВЭД (по сертификату соответствия)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s138">
              <Data s:Type="String">Порядковый номер сертификата соответствия</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="s155">
              <Data s:Type="String">Торговая наценка</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38548268">
              <Data s:Type="String">Рекомендуемая расходная цена</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38051368">
              <Data s:Type="String">Страна изготовитель</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="s138">
              <Data s:Type="String">НДС поставщика</Data>
            </Cell>
            <Cell s:StyleID="s155">
              <Data s:Type="String">Срок реализации, дней</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="m38548492">
              <Data s:Type="String">Поставка через кого (магазин; РЦ)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38548472">
              <Data s:Type="String">Частная марка</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38548452">
              <Data s:Type="String">Формат магазина (-1;0;1;2)</Data>
            </Cell>
            <Cell s:MergeDown="1" s:StyleID="m38548432">
              <Data s:Type="String">Примечание</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="m38548372">
              <Data s:Type="String">матрица</Data>
            </Cell>
            <Cell s:MergeAcross="1" s:StyleID="m38548148">
              <Data s:Type="String">заявка</Data>
            </Cell>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="77.25" s:StyleID="s157">
            <Cell s:Index="19" s:StyleID="s158">
              <Data s:Type="String">высота слоя, м</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s158">
              <Data s:Type="String">количество   ед-ц в слое</Data>
              <NamedCell s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s159">
              <Data s:Type="String">вес БРУТТО единицы товара, кг</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s159">
              <Data s:Type="String">вес НЕТТО единицы товара, кг</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s138">
              <Data s:Type="String">Длина</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s138">
              <Data s:Type="String">Ширина</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s138">
              <Data s:Type="String">Высота</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:Index="28" s:StyleID="s138">
              <Data s:Type="String">имп</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s138">
              <Data s:Type="String">от</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s138">
              <Data s:Type="String">отеч.</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:Index="36" s:StyleID="s138">
              <Data s:Type="Number">10</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s138">
              <Data s:Type="Number">18</Data>
              <NamedCell
      s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s139">
              <NamedCell s:Name="_FilterDatabase"/>
            </Cell>
            <Cell s:StyleID="s131">
              <Data s:Type="String">РЦ</Data>
            </Cell>
            <Cell s:StyleID="s131">
              <Data s:Type="String">М</Data>
            </Cell>
            <Cell s:Index="44" s:StyleID="s154">
              <Data s:Type="String">да</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="String">нет </Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="String">открыть</Data>
            </Cell>
            <Cell s:StyleID="s154">
              <Data s:Type="String">закрыть</Data>
            </Cell>
          </Row>
          <root/>
          <Row s:Hidden="1">
            <Cell s:StyleID="s93"/>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="8" s:StyleID="s92"/>
            <Cell s:Index="10" s:StyleID="s99"/>
            <Cell s:StyleID="s99"/>
            <Cell s:StyleID="s99"/>
            <Cell s:StyleID="s99"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="55.5" s:StyleID="s103">
            <Cell s:StyleID="s135"/>
            <Cell s:StyleID="s104"/>
            <Cell s:StyleID="s140">
              <s:Data s:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40">
                <B>
                  ВСЕ, ЧТО ОТМЕЧЕННО <Font
        html:Color="#FFCC00">ЖЕЛТЫМ</Font><Font> - ОБЯЗАТЕЛЬНО ДЛЯ ЗАПОЛНЕНИЯ!!!!!!!!!!!!!!!!</Font>
                </B>
              </s:Data>
            </Cell>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:Index="39" s:StyleID="s102"/>
          </Row>
          <Row s:AutoFitHeight="0" s:Height="55.5" s:StyleID="s103">
            <Cell s:StyleID="s135"/>
            <Cell s:StyleID="s104"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:Index="39" s:StyleID="s102"/>
          </Row>
          <Row s:StyleID="s103">
            <Cell s:StyleID="s135"/>
            <Cell s:StyleID="s104"/>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s101">
              <Data s:Type="String">Поставщик     </Data>
            </Cell>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s104">
              <Data s:Type="String">Покупатель</Data>
            </Cell>
            <Cell s:StyleID="s102"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:StyleID="s101"/>
            <Cell s:Index="39" s:StyleID="s101"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">Пояснения</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">Столбец №:</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">2</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Пишется &quot;нов&quot; только если товар новый, его нет  Перекрестке и Пятерочке</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">3</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Пишется PLU, если товар уже есть </Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">4</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">ТМ может и не быть. ТМ обязательно на языке оригинала </Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">5</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Наименование без сокращений. Список разрешенных сокращений на листе &quot;сокращения&quot;</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">6</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">финансовый код  (БУКВА И ЦИФРА)</Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">7</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Параметры , если товар весовой, то не заполняется</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">8</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Если мы продаем весом, а вы возите пакетами по 2,5, то вы ставите кг (Цена тоже за кг!!!, а не за пакет)</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">9</Data>
            </Cell>
            <Cell s:StyleID="s118">
              <s:Data s:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40">
                <B>
                  <I>
                    <Font html:Color="#FF0000">ОБЯЗАТЕЛЬНО</Font>
                  </I>
                  <Font> для алкогольной и спиртосодержащей продукции</Font>
                </B>
              </s:Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">10</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">К заполнению ОБЯЗАТЕЛЬНО для алкогольной продукции!</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">11</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">ставится отметка, если товар фасованный</Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122">
              <Data s:Type="String">пят</Data>
            </Cell>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">12</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">ставится отметка, если товар маркированный</Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122">
              <Data s:Type="String">пят</Data>
            </Cell>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">13</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Если ед.измерения- шт, то штрих-код должен быть обязательно, если- кг, то его быть не должно. </Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">14</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Артикул поставщика</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">15</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Наша цена</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">16</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Ценовой сегмент</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">17</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Сколько шт/месте, кг/месте</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">18</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">По сколько минимально кг, шт отгружает поставщик</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">19</Data>
            </Cell>
            <Cell s:StyleID="s118">
              <Data s:Type="String">Заполняется для всех товаров без исключения</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s127">
            <Cell s:StyleID="s136">
              <Data s:Type="String">20</Data>
            </Cell>
            <Cell s:StyleID="s121">
              <Data s:Type="String">Габариты товара, см схему ниже </Data>
            </Cell>
            <Cell s:StyleID="s121"/>
            <Cell s:StyleID="s125"/>
            <Cell s:StyleID="s126"/>
            <Cell s:Index="10" s:StyleID="s129"/>
            <Cell s:StyleID="s129"/>
            <Cell s:StyleID="s129"/>
            <Cell s:StyleID="s129"/>
            <Cell s:Index="24" s:StyleID="s130"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">21</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">количество упаковок на паллете </Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s137">
              <Data s:Type="Number">22</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Именно как называется производитель!!! (ЗАО &quot;Дзержинский Мясокомбинат&quot;, а не просто Дэмка, если товар имп. то на иностранном языке, например &quot;Sea Drift Fish Co, LTD&quot;)</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s137"/>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Не нужно писать - Россия или Белорусия, это не производитель</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="Number">23</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Если произведен в РФ- отеч., если в другой стране (в т.ч. Белорусия и Украина)- имп.</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">24</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Код ОКПД - это 6 цифр, не больше и не меньше. Его можно найти в сертификате</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">25</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Можно писать №РОСС RU.ПН 33. В01324, можно только последние цифры В01324</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">26</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Стандартная торговая наценка </Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">27</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">рекомендуемая расходная цена</Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">28</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">страна изготовитель </Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">29</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">НДС</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">30</Data>
            </Cell>
            <Cell s:StyleID="s119">
              <Data s:Type="String">Срок реализации - дней (не часы, не месяцы, а именно дни)</Data>
            </Cell>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s119"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:StyleID="s120"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:Index="39" s:StyleID="s120"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">31</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">отметка,через кого поставка </Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">32</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">частная марка </Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">33</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">формат магазина (-1;0;1;2)</Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="String">34</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">примечание </Data>
            </Cell>
            <Cell s:Index="5" s:StyleID="s125"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75" s:StyleID="s121">
            <Cell s:StyleID="s136">
              <Data s:Type="Number">35</Data>
            </Cell>
            <Cell s:StyleID="s124">
              <Data s:Type="String">матрица</Data>
            </Cell>
            <Cell s:StyleID="s127"/>
            <Cell s:StyleID="s128"/>
            <Cell s:StyleID="s128"/>
            <Cell s:Index="10" s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:StyleID="s122"/>
            <Cell s:Index="24" s:StyleID="s123"/>
            <Cell s:Index="32" s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
            <Cell s:StyleID="s124"/>
          </Row>
          <Row s:Height="15.75">
            <Cell s:StyleID="s136">
              <Data s:Type="Number">36</Data>
            </Cell>
            <Cell>
              <Data s:Type="String">заявка</Data>
            </Cell>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:StyleID="s116">
              <Data s:Type="String">длина</Data>
            </Cell>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:StyleID="s91">
              <Data s:Type="String">высота</Data>
            </Cell>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="3" s:MergeAcross="1" s:StyleID="s270">
              <Data s:Type="String">Лицевая сторона товара при размещении на полке</Data>
            </Cell>
            <Cell s:Index="8" s:StyleID="s115">
              <Data s:Type="String">ширина</Data>
            </Cell>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
            <Cell s:Index="24" s:StyleID="s92"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
          <Row>
            <Cell s:Index="4" s:StyleID="s94"/>
          </Row>
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
          <PageSetup>
            <Layout x:Orientation="Landscape"/>
            <Header x:Margin="0.28000000000000003"/>
            <Footer x:Margin="0.17"/>
            <PageMargins x:Bottom="0.17" x:Left="0.19685039370078741"
             x:Right="0.19685039370078741" x:Top="0.39370078740157483"/>
          </PageSetup>
          <FitToPage/>
          <Print>
            <ValidPrinterInfo/>
            <PaperSizeIndex>9</PaperSizeIndex>
            <Scale>31</Scale>
            <HorizontalResolution>1200</HorizontalResolution>
            <VerticalResolution>1200</VerticalResolution>
          </Print>
          <Zoom>80</Zoom>
          <PageBreakZoom>60</PageBreakZoom>
          <Selected/>
          <Panes>
            <Pane>
              <Number>3</Number>
              <ActiveRow>15</ActiveRow>
              <ActiveCol>18</ActiveCol>
            </Pane>
          </Panes>
          <ProtectObjects>False</ProtectObjects>
          <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
      </Worksheet>
    </Workbook>
  </xsl:template>
</xsl:stylesheet>
