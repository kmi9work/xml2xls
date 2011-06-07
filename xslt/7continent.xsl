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
          <Author>misenko</Author>
          <LastAuthor>Rogozhin Anton Yurevich</LastAuthor>
          <Created>2004-04-14T08:56:53Z</Created>
          <LastSaved>2009-10-18T11:42:10Z</LastSaved>
          <Company>JSV The Seventh Continent</Company>
          <Version>12.00</Version>
        </DocumentProperties>
        <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
          <WindowHeight>8835</WindowHeight>
          <WindowWidth>9180</WindowWidth>
          <WindowTopX>120</WindowTopX>
          <WindowTopY>120</WindowTopY>
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
          <Style ss:ID="s63" ss:Name="Гиперссылка">
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Color="#0000FF"
             ss:Underline="Single"/>
          </Style>
          <Style ss:ID="s20" ss:Name="Процентный">
            <NumberFormat ss:Format="0%"/>
          </Style>
          <Style ss:ID="s16" ss:Name="Финансовый">
            <NumberFormat
             ss:Format="_-* #,##0.00_р_._-;\-* #,##0.00_р_._-;_-* &quot;-&quot;??_р_._-;_-@_-"/>
          </Style>
          <Style ss:ID="m38878780">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="12" ss:Bold="1"/>
            <Interior/>
          </Style>
          <Style ss:ID="m38878496">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="9"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878516">
            <Alignment ss:Vertical="Justify" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
          </Style>
          <Style ss:ID="m38878536">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="9"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878556">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878272">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878292">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38878312">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="9"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878332">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878352">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="9"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878372">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878048">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878068">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878088">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878108">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38878128">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38878148">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38877844">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877864">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="m38877884">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877904">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877924">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877944">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877600">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877620">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38877640">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877660">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38877680">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877700">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877376">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877396">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877416">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877436">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38877456">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877476">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38877152">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877172">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38877192">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877212">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877232">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877252">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38876948">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38876968">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38876988">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877008">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38877028">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="m38877048">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38876704">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="m38876724">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="m38876744">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
          </Style>
          <Style ss:ID="m38876764">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
            <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="m38876784">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38876804">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38876480">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
          </Style>
          <Style ss:ID="m38876500">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="m38876520">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="m38876540">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="m38876560">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="m38876580">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
          </Style>
          <Style ss:ID="m38876600">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="m38876620">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="m38876640">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38876660">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38876256">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="m38876276">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876296" ss:Parent="s63">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Protection/>
          </Style>
          <Style ss:ID="m38876316">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38876336">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38876356">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8"/>
          </Style>
          <Style ss:ID="m38876376">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38876396">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="m38876416">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s64">
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s65">
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s66">
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s67">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="12"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s68">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s69">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s70">
            <Alignment ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s71">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s72">
            <Alignment ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s73">
            <Alignment ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s74">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s75">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="24"/>
          </Style>
          <Style ss:ID="s76">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="12"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s77">
            <Alignment ss:Vertical="Bottom"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="12"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s96">
            <Alignment ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s97">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s98">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s99">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s100">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s101">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s102">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s103">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s104">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s105">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s107" ss:Parent="s63">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Protection/>
          </Style>
          <Style ss:ID="s108" ss:Parent="s63">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Protection/>
          </Style>
          <Style ss:ID="s109">
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s110">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
            <Interior/>
          </Style>
          <Style ss:ID="s174">
            <Alignment ss:Vertical="Top"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s175">
            <Alignment ss:Horizontal="Center" ss:Vertical="Top"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s191">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
          </Style>
          <Style ss:ID="s192">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
          </Style>
          <Style ss:ID="s193">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
          </Style>
          <Style ss:ID="s216">
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Color="#FF0000"/>
          </Style>
          <Style ss:ID="s217">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Italic="1"/>
          </Style>
          <Style ss:ID="s219">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="24"/>
          </Style>
          <Style ss:ID="s221">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="8" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s222">
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Color="#FF0000" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s263">
            <Borders>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Color="#FF0000" ss:Bold="1"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s317">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="9"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s321">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="9"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s323">
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s326">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="14"
             ss:Color="#FF0000" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s329">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="14"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s330">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Size="14"/>
          </Style>
          <Style ss:ID="s331">
            <Borders/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s333">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Size="14" ss:Color="#FF0000"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s334">
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Color="#0000FF"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s344">
            <Alignment ss:Vertical="Center" ss:WrapText="1"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="12" ss:Bold="1"/>
            <Interior/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s345">
            <Alignment ss:Vertical="Center" ss:WrapText="1"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:Family="Swiss" ss:Size="12" ss:Bold="1"/>
            <Interior/>
          </Style>
          <Style ss:ID="s347" ss:Parent="s20">
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s348">
            <Font ss:FontName="Arial"/>
            <NumberFormat ss:Format="0.000"/>
          </Style>
          <Style ss:ID="s349">
            <Font ss:FontName="Arial"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s350">
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
          </Style>
          <Style ss:ID="s351">
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss" ss:Color="#3366FF"/>
            <Interior/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s353">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
            <Interior/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s356">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" ss:Bold="1"/>
            <Interior/>
            <NumberFormat ss:Format="Short Date"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s358">
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s359">
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="0.000"/>
          </Style>
          <Style ss:ID="s360">
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s361">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s364">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
            <NumberFormat ss:Format="0"/>
            <Protection/>
          </Style>
          <Style ss:ID="s366">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s368">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s370" ss:Parent="s63">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s371">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
            <NumberFormat ss:Format="Short Date"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s373">
            <Borders/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s374">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s375">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s376">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s378" ss:Parent="s16">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat/>
          </Style>
          <Style ss:ID="s379">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s380">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s381">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s382" ss:Parent="s20">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s383">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s384">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s385">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s386">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s387">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s388">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s389">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
          </Style>
          <Style ss:ID="s390" ss:Parent="s20">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s391">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s392" ss:Parent="s20">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s393">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s394">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s395">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s396">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s397">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s398">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s399">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          <Interior ss:Color="#888888" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s1001">
            <Borders/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s1002">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s1003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s1004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s1005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s1006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s1007">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
            <Interior/>
          </Style>
          <Style ss:ID="s1008" ss:Parent="s20">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s1010" ss:Parent="s20">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1011">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s1012">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s1013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s1014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s1015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1016">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s2001">
            <Borders/>
            <Font ss:FontName="Arial"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s2002">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s2003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s2004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s2005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s2006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s2007">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s2008" ss:Parent="s20">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s2010" ss:Parent="s20">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2011">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s2012">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s2013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s2014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s2015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2016">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s3001">
            <Borders/>
            <Font ss:FontName="Arial"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s3002">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s3003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s3004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s3005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s3006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s3007">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s3008" ss:Parent="s20">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s3010" ss:Parent="s20">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3011">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s3012">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s3013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s3014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s3015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3016">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s4001">
            <Borders/>
            <Font ss:FontName="Arial"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s4002">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s4003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s4004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s4005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
            <NumberFormat ss:Format="0"/>
          </Style>
          <Style ss:ID="s4006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s4007">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial Cyr" x:CharSet="204" x:Family="Swiss"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s4008" ss:Parent="s20">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s4010" ss:Parent="s20">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4011">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
            <NumberFormat ss:Format="Fixed"/>
          </Style>
          <Style ss:ID="s4012">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s4013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
          <Style ss:ID="s4014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s4015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Tahoma" x:CharSet="204" x:Family="Swiss"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4016">
            <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial"/>
          </Style>
        </Styles>
        <Worksheet ss:Name="Комментарий">
          <Table ss:ExpandedColumnCount="38" x:FullColumns="1"
           x:FullRows="1" ss:StyleID="s64" ss:DefaultRowHeight="31.5">
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="23.25"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="52.5"/>
            <Column ss:StyleID="s64" ss:Width="74.25"/>
            <Column ss:StyleID="s64" ss:Hidden="1" ss:AutoFitWidth="0" ss:Width="102.75"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="133.5"/>
            <Column ss:StyleID="s64" ss:Width="174"/>
            <Column ss:StyleID="s64" ss:Width="119.25"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="97.5"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="114.75"/>
            <Column ss:StyleID="s64" ss:Width="115.5"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="93"/>
            <Column ss:Index="22" ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="109.5"/>
            <Column ss:StyleID="s64" ss:Width="88.5" ss:Span="1"/>
            <Column ss:Index="25" ss:StyleID="s64" ss:Width="45.75"/>
            <Column ss:StyleID="s64" ss:Width="84"/>
            <Column ss:Index="33" ss:StyleID="s64" ss:Hidden="1" ss:AutoFitWidth="0"
             ss:Span="1"/>
            <Column ss:Index="35" ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="130.5"/>
            <Column ss:Index="37" ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="58.5"/>
            <Row ss:AutoFitHeight="0"/>
            <Row ss:AutoFitHeight="0" ss:Height="15.75">
              <Cell ss:Index="2" ss:StyleID="s65"/>
              <Cell ss:StyleID="s334"/>
              <Cell ss:MergeAcross="6" ss:MergeDown="1" ss:StyleID="m38878780">
                <Data
      ss:Type="String">Ценовой лист</Data>
              </Cell>
              <Cell ss:StyleID="s344"/>
              <Cell ss:StyleID="s345"/>
              <Cell ss:StyleID="s347"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:Index="20" ss:StyleID="s348"/>
              <Cell ss:Index="24" ss:StyleID="s349"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="15.75">
              <Cell ss:Index="2" ss:StyleID="s65"/>
              <Cell ss:StyleID="s334"/>
              <Cell ss:Index="11" ss:StyleID="s344"/>
              <Cell ss:StyleID="s345"/>
              <Cell ss:StyleID="s347"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:Index="20" ss:StyleID="s348"/>
              <Cell ss:Index="24" ss:StyleID="s349"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="15.75">
              <Cell ss:Index="2" ss:StyleID="s350"/>
              <Cell ss:StyleID="s351"/>
              <Cell ss:MergeAcross="2" ss:StyleID="s353">
                <Data ss:Type="String">Дата вступления в силу:</Data>
              </Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s356">
                <Data ss:Type="DateTime">2005-09-15T00:00:00.000</Data>
              </Cell>
              <Cell ss:StyleID="s358"/>
              <Cell ss:StyleID="s350"/>
              <Cell ss:StyleID="s347"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s359"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s360"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="15.75">
              <Cell ss:Index="2" ss:StyleID="s350"/>
              <Cell ss:StyleID="s351"/>
              <Cell ss:MergeAcross="2" ss:StyleID="s361">
                <Data ss:Type="String">Покупатель:</Data>
              </Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s356">
                <Data ss:Type="String">ОАО &quot;Седьмой континент&quot;</Data>
              </Cell>
              <Cell ss:StyleID="s358"/>
              <Cell ss:StyleID="s350"/>
              <Cell ss:StyleID="s347"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s359"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s360"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="15.75">
              <Cell ss:Index="2" ss:StyleID="s350"/>
              <Cell ss:StyleID="s351"/>
              <Cell ss:MergeAcross="2" ss:StyleID="s361">
                <Data ss:Type="String">Дата формирования файла:</Data>
              </Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s371">
                <Data ss:Type="DateTime">2005-08-03T00:00:00.000</Data>
              </Cell>
              <Cell ss:StyleID="s358"/>
              <Cell ss:StyleID="s350"/>
              <Cell ss:StyleID="s347"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s359"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s360"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="15.75">
              <Cell ss:Index="2" ss:StyleID="s350"/>
              <Cell ss:StyleID="s351"/>
              <Cell ss:StyleID="s361"/>
              <Cell ss:StyleID="s361"/>
              <Cell ss:StyleID="s361"/>
              <Cell ss:StyleID="s371"/>
              <Cell ss:StyleID="s371"/>
              <Cell ss:StyleID="s371"/>
              <Cell ss:StyleID="s371"/>
              <Cell ss:StyleID="s358"/>
              <Cell ss:StyleID="s350"/>
              <Cell ss:StyleID="s347"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s359"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s360"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
              <Cell ss:StyleID="s65"/>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="15" ss:StyleID="s373">
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">1</Data>
              </Cell>
              <Cell ss:StyleID="s375">
                <Data ss:Type="Number">2</Data>
              </Cell>
              <Cell ss:StyleID="s376">
                <Data ss:Type="Number">3</Data>
              </Cell>
              <Cell ss:StyleID="s376">
                <Data ss:Type="Number">4</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">5</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">6</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">7</Data>
              </Cell>
              <Cell ss:StyleID="s378">
                <Data ss:Type="Number">8</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">9</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">10</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">11</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">12</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">13</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">14</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">15</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">16</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">17</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">18</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">19</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">20</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">21</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">22</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">23</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">24</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">25</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">26</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">25</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">24</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">25</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">26</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">27</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">28</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">25</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">25</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">29</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">30</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">31</Data>
              </Cell>
              <Cell ss:StyleID="s374">
                <Data ss:Type="Number">32</Data>
              </Cell>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="78" ss:StyleID="s373">
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">№ п.п.</Data>
              </Cell>
              <Cell ss:StyleID="s380">
                <Data ss:Type="String">Состояние товара (новый)</Data>
              </Cell>
              <Cell ss:StyleID="s381">
                <Data ss:Type="String">Штрих-код</Data>
              </Cell>
              <Cell ss:StyleID="s381">
                <Data ss:Type="String">Артикул</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Группировка товаров Поставщика</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Наименование</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Торговая марка</Data>
              </Cell>
              <Cell ss:StyleID="s382">
                <Data ss:Type="String">НДС (20% или 18% или 10% или  0%)</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Цена с НДС</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Цена для ГиперМаркета</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Цена для Санкт-Петербурга</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">Цена для Рязани</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">Цена для Челябинска</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">Цена для Красноярска</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">Цена для Уфы</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Валюта  (RUR, EUR, USD)</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">БАЗА ДАННЫХ / НДС /</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">БАЗА ДАННЫХ /Цена с НДС/</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">БАЗА ДАННЫХ /Цена для Гипермаркета/</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">БАЗА ДАННЫХ /Цена для Санкт-Петербурга/</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">БАЗА ДАННЫХ /Цена для Рязани/</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">БАЗА ДАННЫХ /Цена для Челябинска/</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">БАЗА ДАННЫХ /Цена для Красноярска/</Data>
              </Cell>
              <Cell ss:StyleID="s383">
                <Data ss:Type="String">БАЗА ДАННЫХ /Цена для Уфы/</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">БАЗА ДАННЫХ /Валюта/</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Производитель</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Страна происхождения</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Весовой или штучный(шт, кг)</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Для штучного товара Кол-во в ед (0,75 или 1)</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Для штучного товара Ед.изм (шт, л, кг)</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Кол-во штук в коробке  (шт)</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Мин кол-во заказываемого товара (кол-во шт)</Data>
              </Cell>
              <Cell ss:StyleID="s384">
                <Data ss:Type="String">AssortmentId</Data>
              </Cell>
              <Cell ss:StyleID="s384">
                <Data ss:Type="String">AssortmentId</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Грузовая таможенная декларация</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Код 005-93</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Код ВЭД</Data>
              </Cell>
              <Cell ss:StyleID="s379">
                <Data ss:Type="String">Ценовая категория(премиум, средняя, ординар)</Data>
              </Cell>
            </Row>
            <root />
          </Table>
          <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
            <Unsynced/>
            <Print>
              <ValidPrinterInfo/>
              <PaperSizeIndex>9</PaperSizeIndex>
              <HorizontalResolution>600</HorizontalResolution>
              <VerticalResolution>600</VerticalResolution>
            </Print>
            <Selected/>
            <TopRowVisible>1</TopRowVisible>
            <LeftColumnVisible>1</LeftColumnVisible>
            <Panes>
              <Pane>
                <Number>3</Number>
                <ActiveRow>83</ActiveRow>
                <ActiveCol>26</ActiveCol>
                <RangeSelection>C27</RangeSelection>
              </Pane>
            </Panes>
            <ProtectObjects>False</ProtectObjects>
            <ProtectScenarios>False</ProtectScenarios>
          </WorksheetOptions>
          <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
            <Range>R93C4</Range>
            <Type>Whole</Type>
            <Qualifier>Greater</Qualifier>
            <UseBlank/>
            <Value>1</Value>
            <InputTitle>Код продукта</InputTitle>
            <ErrorMessage>Код продукта цифровое поле.</ErrorMessage>
            <ErrorTitle>Ошибка!!!</ErrorTitle>
          </DataValidation>
        </Worksheet>
      </Workbook>
    </xsl:template>
</xsl:stylesheet>
