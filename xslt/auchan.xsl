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
          <Author>RU00110623</Author>
          <LastAuthor>Rogozhin Anton Yurevich</LastAuthor>
          <LastPrinted>2006-09-15T15:01:34Z</LastPrinted>
          <Created>2006-03-20T13:49:03Z</Created>
          <LastSaved>2009-10-18T12:08:15Z</LastSaved>
          <Company>Auchan</Company>
          <Version>12.00</Version>
        </DocumentProperties>
        <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
          <WindowHeight>4560</WindowHeight>
          <WindowWidth>19170</WindowWidth>
          <WindowTopX>-30</WindowTopX>
          <WindowTopY>4245</WindowTopY>
          <TabRatio>157</TabRatio>
          <ProtectStructure>False</ProtectStructure>
          <ProtectWindows>False</ProtectWindows>
        </ExcelWorkbook>
        <Styles>
          <Style ss:ID="Default" ss:Name="Normal">
            <Alignment ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:CharSet="204"/>
            <Interior/>
            <NumberFormat/>
            <Protection/>
          </Style>
          <Style ss:ID="m38876032">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"/>
          </Style>
          <Style ss:ID="m38875808">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Color="#993300" ss:Bold="1"/>
            <NumberFormat ss:Format="0%"/>
          </Style>
          <Style ss:ID="m38875908">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="m38875948">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="m38875584">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Color="#993300"/>
          </Style>
          <Style ss:ID="m38875604">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"/>
          </Style>
          <Style ss:ID="m38875380">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          </Style>
          <Style ss:ID="m38875400">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Color="#333333" ss:Bold="1"/>
            <Interior ss:Color="#FF99CC" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="m38875480">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:ShrinkToFit="1"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
          </Style>
          <Style ss:ID="s64">
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s65">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s66">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
            <NumberFormat ss:Format="0%"/>
          </Style>
          <Style ss:ID="s67">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s68">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s69">
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
          </Style>
          <Style ss:ID="s70">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s71">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
            <NumberFormat ss:Format="0%"/>
          </Style>
          <Style ss:ID="s72">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s73">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s74">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
          </Style>
          <Style ss:ID="s75">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
          </Style>
          <Style ss:ID="s76">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s77">
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
          </Style>
          <Style ss:ID="s78">
            <Alignment ss:Vertical="Bottom"/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s79">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s80">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s81">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s82">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s83">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:ShrinkToFit="1"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9"/>
          </Style>
          <Style ss:ID="s84">
            <Alignment ss:Horizontal="Center" ss:Vertical="Top" ss:ShrinkToFit="1"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s85">
            <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Color="#FFFFFF"/>
          </Style>
          <Style ss:ID="s86">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s87">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"/>
          </Style>
          <Style ss:ID="s88">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s89">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s90">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s91">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s92">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s93">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s94">
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
          </Style>
          <Style ss:ID="s95">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
          </Style>
          <Style ss:ID="s96">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s97">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s98">
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
          </Style>
          <Style ss:ID="s99">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
          </Style>
          <Style ss:ID="s100">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <NumberFormat ss:Format="0%"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s101">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s102">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s103">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s104">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s105">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s106">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s107">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat/>
            <Protection/>
          </Style>
          <Style ss:ID="s108">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat/>
            <Protection/>
          </Style>
          <Style ss:ID="s109">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s110">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s111">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
          </Style>
          <Style ss:ID="s112">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s113">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11"/>
            <NumberFormat ss:Format="0%"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s140">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s141">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <NumberFormat/>
            <Protection/>
          </Style>
          <Style ss:ID="s142">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s143">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s144">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <NumberFormat ss:Format="0%"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s145">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s146">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s147">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <NumberFormat/>
            <Protection/>
          </Style>
          <Style ss:ID="s148">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="12"
             ss:Bold="1"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s149">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
          </Style>
          <Style ss:ID="s150">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="12"
             ss:Bold="1"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s151">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s152">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s153">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s154">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s155">
            <Alignment ss:Horizontal="Center" ss:Vertical="Top" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s161">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
          <Style ss:ID="s164">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s165">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss"/>
            <NumberFormat ss:Format="@"/>
          </Style>
          <Style ss:ID="s172">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Bold="1"/>
          </Style>
          <Style ss:ID="s173">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Color="#993300" ss:Bold="1"/>
          </Style>
          <Style ss:ID="s177">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"
             ss:Bold="1"/>
            <Interior ss:Color="#FFCC99" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s180">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"/>
          </Style>
        <Style ss:ID="s181">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="11"/>
            <Interior ss:Color="#888888" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s1001">
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
          </Style>
          <Style ss:ID="s1002">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
          </Style>
          <Style ss:ID="s1005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/> 
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1007">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1008">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1010">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat/>
            <Protection/>
          </Style> 
          <Style ss:ID="s1011">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1012">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
	        <Style ss:ID="s1013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <NumberFormat ss:Format="0%"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s1016">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2001">
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s2002">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s2005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2007">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2008">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2010">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <NumberFormat/>
            <Protection/>
          </Style> 
          <Style ss:ID="s2011">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2012">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
	        <Style ss:ID="s2013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="0%"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s2016">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
                  <Style ss:ID="s3001">
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s3002">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s3005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3007">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3008">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3010">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <NumberFormat/>
            <Protection/>
          </Style> 
          <Style ss:ID="s3011">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3012">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
	        <Style ss:ID="s3013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="0%"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s3016">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
                  <Style ss:ID="s4001">
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s4002">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4003">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4004">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s4005">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4006">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4007">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4008">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4009">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4010">
            <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <NumberFormat/>
            <Protection/>
          </Style> 
          <Style ss:ID="s4011">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4012">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
	        <Style ss:ID="s4013">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="@"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4014">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <NumberFormat ss:Format="0%"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4015">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
          <Style ss:ID="s4016">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
            <Protection ss:Protected="0"/>
          </Style>
        </Styles>
        <Worksheet ss:Name="ОБЩАЯ">
          <Names>
            <NamedRange ss:Name="Print_Titles" ss:RefersTo="=ОБЩАЯ!C1:C3,ОБЩАЯ!R5:R7"/>
            <NamedRange ss:Name="Print_Area" ss:RefersTo="=ОБЩАЯ!R1C1:R30C44"/>
          </Names>
          <Table ss:ExpandedColumnCount="252" x:FullColumns="1"
           x:FullRows="1" ss:StyleID="s69" ss:DefaultRowHeight="14.25">
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="53.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="68.25"/>
            <Column ss:StyleID="s65" ss:Width="432"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="83.25"/>
            <Column ss:StyleID="s65" ss:Width="117"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="47.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="56.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="46.5" ss:Span="1"/>
            <Column ss:Index="10" ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="112.5"/>
            <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="65.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="79.5"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="67.5"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="56.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="42.75"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="40.5"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="45"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="30.75"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="50"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="50"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="50"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="27"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="27.75"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="24"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="24.75" ss:Span="3"/>
            <Column ss:Index="29" ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="29.25"
             ss:Span="3"/>
            <Column ss:Index="33" ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="30.75"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="40.5"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="79.5"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="50.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="58.5"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="74.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="83.25"/>
            <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="112.5"/>
            <Column ss:StyleID="s64" ss:Width="357"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="51.75"/>
            <Column ss:StyleID="s78" ss:AutoFitWidth="0" ss:Width="54"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="45"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="45.75"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="51.75"/>
            <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="279.75"/>
            <Column ss:StyleID="s69" ss:AutoFitWidth="0" ss:Width="82.5" ss:Span="28"/>
            <Row ss:AutoFitHeight="0" ss:Height="27">
              <Cell ss:StyleID="s86">
                <Data ss:Type="String">№ поставщика</Data>
                <NamedCell
      ss:Name="Print_Area"/>
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s64">
                <NamedCell ss:Name="Print_Area"/>
                <NamedCell
      ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s148">
                <NamedCell ss:Name="Print_Area"/>
                <NamedCell
      ss:Name="Print_Titles"/>
              </Cell>
            </Row>
            <Row ss:Height="15">
              <Cell ss:StyleID="s67">
                <NamedCell ss:Name="Print_Area"/>
                <NamedCell
      ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s64">
                <NamedCell ss:Name="Print_Area"/>
                <NamedCell
      ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s149">
                <NamedCell ss:Name="Print_Area"/>
                <NamedCell
      ss:Name="Print_Titles"/>
              </Cell>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="23.25">
              <Cell ss:StyleID="s86">
                <Data ss:Type="String">Название поставщика</Data>
                <NamedCell
      ss:Name="Print_Area"/>
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s64">
                <NamedCell ss:Name="Print_Area"/>
                <NamedCell
      ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s150">
                <NamedCell ss:Name="Print_Area"/>
                <NamedCell
      ss:Name="Print_Titles"/>
              </Cell>
            </Row>
            <Row ss:Height="15">
              <Cell ss:Index="39" ss:MergeAcross="5" ss:StyleID="m38875400">
                <Data
      ss:Type="String">К ВНИМАНИЮ ПОСТАВЩИКА ! Раздел заполняется только закупщиком</Data>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="41.25" ss:StyleID="s72">
              <Cell ss:MergeDown="2" ss:StyleID="s177">
                <Data ss:Type="String">№ п/п</Data>
                <NamedCell
      ss:Name="Print_Area"/>
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="s177">
                <Data ss:Type="String">КОД ТОВАРА</Data>
                <NamedCell
      ss:Name="Print_Area"/>
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="s177">
                <Data ss:Type="String">Наименование товара / для заказа. Обязательно указать вес/объем/длина</Data>
                <NamedCell
      ss:Name="Print_Area"/>
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="s164">
                <Data ss:Type="String">внутренн.код товара у пост </Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="s164">
                <Data ss:Type="String">ш-к</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s164">
                <Data ss:Type="String">номенклатура</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="s173">
                <Data ss:Type="String">цена закупки базовая без НДС </Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="m38875808">
                <Data ss:Type="String">скидка по счету</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="m38875584">
                <Data ss:Type="String">ставка НДС</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="m38876032">
                <Data ss:Type="String">кол-во в коробке</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="2" ss:StyleID="m38875604">
                <Data ss:Type="String">срок реализации</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="s180">
                <Data ss:Type="String">тип товара</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s161">
                <Data ss:Type="String">№ колл </Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="s161"><Data
      ss:Type="String">ИНФОРМАЦИЯ О РАЗМЕРЕ ТОВАРА В САНТИМЕТРАХ</Data><NamedCell
      ss:Name="Print_Titles"/><NamedCell ss:Name="Print_Area"/></Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s161">
                <Data ss:Type="String">единица закупки</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s161">
                <Data ss:Type="String">единица продажи</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s161">
                <Data ss:Type="String">хар-ка товара </Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="s161">
                <Data ss:Type="String">происхожд. товара</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s161">
                <Data ss:Type="String">только период промо</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38875380">
                <Data ss:Type="String">Ветеринар.сертиф.         ДА / НЕТ                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             </Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeAcross="3" ss:StyleID="s155">
                <Data ss:Type="String">Спиртосодержащая продукция более 1.5%&#10;(за искл. пива и шоколада)</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s164">
                <Data ss:Type="String">описание на ценнике</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s161">
                <Data ss:Type="String">Пост. НАЦИОНАЛЬНЫЙ</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s165">
                <Data ss:Type="String">МО</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s161">
                <Data ss:Type="String">СП</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38875948">
                <Data ss:Type="String">ЕК</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38875908">
                <Data ss:Type="String">НН</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38875480">
                <Data ss:Type="String">Если необходимо, укажите название магазина/-ов</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
            </Row>
            <Row ss:Height="51" ss:StyleID="s72">
              <Cell ss:Index="6" ss:MergeDown="1" ss:StyleID="s172">
                <Data ss:Type="String">рын.</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s172">
                <Data ss:Type="String">сегм.</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s172">
                <Data ss:Type="String">катег.</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="s164">
                <Data ss:Type="String">сем.</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:Index="15" ss:StyleID="s87">
                <Data ss:Type="String">ФК </Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s87">
                <Data ss:Type="String">ИК</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:Index="21" ss:StyleID="s68">
                <Data ss:Type="String">шт</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">кг</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">л</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">м</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">шт</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">кг</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">л</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">м</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">1Ц</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">СМА</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">1 ЦПА</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">НМ</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">Pосс.</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s68">
                <Data ss:Type="String">Др.</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:Index="37" ss:StyleID="s84">
                <Data ss:Type="String">содержит спирт: &#10; ДА / НЕТ&#10;</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s84">
                <Data ss:Type="String">% содержания спирта&#10;</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s84">
                <Data ss:Type="String">код декларации&#10;для этилового спирта</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s84">
                <Data ss:Type="String">объем в литрах единицы закупки&#10;</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:Index="48" ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="21" ss:StyleID="s74">
              <Cell ss:Index="15" ss:StyleID="s88">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s88">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s88">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s88"><Data ss:Type="String">глубина</Data><NamedCell ss:Name="Print_Titles"/><NamedCell ss:Name="Print_Area"/></Cell>
              <Cell ss:StyleID="s88"><Data ss:Type="String">ширина</Data><NamedCell ss:Name="Print_Titles"/><NamedCell ss:Name="Print_Area"/></Cell>
              <Cell ss:StyleID="s88"><Data ss:Type="String">высота</Data><NamedCell ss:Name="Print_Titles"/><NamedCell ss:Name="Print_Area"/></Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s83">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s73">
                <NamedCell ss:Name="Print_Titles"/>
                <NamedCell
      ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:Index="42" ss:StyleID="s75">
                <Data ss:Type="Number">801</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s79">
                <Data ss:Type="String">000</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s75">
                <Data ss:Type="Number">100</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s75">
                <Data ss:Type="Number">200</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s75">
                <Data ss:Type="Number">300</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s75">
                <Data ss:Type="String">001, 002, ... 007 ...</Data>
                <NamedCell
      ss:Name="Print_Titles"/>
                <NamedCell ss:Name="Print_Area"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
              <Cell ss:StyleID="s69">
                <NamedCell ss:Name="Print_Titles"/>
              </Cell>
            </Row>
            <root />
          </Table>
          <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
            <PageSetup>
              <Layout x:Orientation="Landscape" x:CenterHorizontal="1"/>
              <Header x:Margin="0.22"/>
              <Footer x:Margin="0.27559055118110237"
               x:Data="&amp;LУтверждено: Поставщик ___________&amp;C                            Ашан: Руководитель направления___________                                                                          Руководитель отдела___________              &amp;RДиректор ЦОЗ___________     "/>
              <PageMargins x:Bottom="0.51181102362204722" x:Left="0.14000000000000001"
               x:Right="0.11811023622047245" x:Top="0.59055118110236227"/>
            </PageSetup>
            <Print>
              <FitWidth>0</FitWidth>
              <ValidPrinterInfo/>
              <PaperSizeIndex>9</PaperSizeIndex>
              <Scale>60</Scale>
              <HorizontalResolution>600</HorizontalResolution>
              <VerticalResolution>600</VerticalResolution>
            </Print>
            <Zoom>75</Zoom>
            <PageBreakZoom>60</PageBreakZoom>
            <Selected/>
            <Panes>
              <Pane>
                <Number>3</Number>
                <ActiveRow>7</ActiveRow>
                <ActiveCol>32</ActiveCol>
                <RangeSelection>R8C33:R9C33</RangeSelection>
              </Pane>
            </Panes>
            <ProtectObjects>False</ProtectObjects>
            <ProtectScenarios>False</ProtectScenarios>
          </WorksheetOptions>
          <PageBreaks xmlns="urn:schemas-microsoft-com:office:excel">
            <ColBreaks>
              <ColBreak>
                <Column>16</Column>
                <RowEnd>29</RowEnd>
              </ColBreak>
              <ColBreak>
                <Column>37</Column>
                <RowEnd>29</RowEnd>
              </ColBreak>
            </ColBreaks>
          </PageBreaks>
          <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
            <Range>R11C38:R27C38,R11C3:R28C3</Range>
            <Type>TextLength</Type>
            <Min>0</Min>
            <Max>30</Max>
            <ErrorMessage>Название товара не более 30 символов.</ErrorMessage>
          </DataValidation>
          <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
            <Range>R7C32:R28C37</Range>
            <Type>TextLength</Type>
            <Min>0</Min>
            <Max>60</Max>
            <ErrorMessage>Описание на ценнике не более 60 символов.</ErrorMessage>
          </DataValidation>
        </Worksheet>
      </Workbook>
    </xsl:template>
</xsl:stylesheet>
