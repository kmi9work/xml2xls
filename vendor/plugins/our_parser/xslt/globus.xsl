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
          <Font ss:FontName="MS Sans Serif"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style ss:ID="s21" ss:Name="Обычный 2">
          <Alignment ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif"/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style ss:ID="s23" ss:Parent="s21">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s26" ss:Parent="s21">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#CCFFFF" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s27" ss:Parent="s21">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s29" ss:Parent="s21">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s30" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s31" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#CCFFFF" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s32" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="\ #############\ 0000000000000"/>
        </Style>
        <Style ss:ID="s33" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#CCFFFF" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s34" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s36" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s37" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s38" ss:Parent="s21">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior/>
          <NumberFormat ss:Format="0000000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s39" ss:Parent="s21">
          <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s40" ss:Parent="s21">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s41" ss:Parent="s21">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="#######"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s42" ss:Parent="s21">
          <Borders/>
          <Interior/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s43" ss:Parent="s21">
          <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
          <Borders/>
          <Interior/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s44" ss:Parent="s21">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s45" ss:Parent="s21">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="000#"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s47" ss:Parent="s21">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="####\0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s48" ss:Parent="s21">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="#########\000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s49" ss:Parent="s21">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="########00000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s50" ss:Parent="s21">
          <Borders/>
          <Interior/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s51" ss:Parent="s21">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="Fixed"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s52" ss:Parent="s21">
          <Borders/>
          <Interior/>
          <NumberFormat ss:Format="Short Date"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s53" ss:Parent="s21">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="####0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s54" ss:Parent="s21">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior/>
          <NumberFormat ss:Format="0.0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s55" ss:Parent="s21">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s121" ss:Name="Обычный 3">
          <Alignment ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style ss:ID="s123" ss:Parent="s121">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s126" ss:Parent="s121">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s127" ss:Parent="s121">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s129" ss:Parent="s121">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s130" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s131" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s132" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="\ #############\ 0000000000000"/>
        </Style>
        <Style ss:ID="s133" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s134" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s136" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s137" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s138" ss:Parent="s121">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0000000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s139" ss:Parent="s121">
          <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s140" ss:Parent="s121">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s141" ss:Parent="s121">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="#######"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s142" ss:Parent="s121">
          <Borders/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s143" ss:Parent="s121">
          <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
          <Borders/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s144" ss:Parent="s121">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s145" ss:Parent="s121">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="000#"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s147" ss:Parent="s121">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="####\0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s148" ss:Parent="s121">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="#########\000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s149" ss:Parent="s121">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="########00000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s150" ss:Parent="s21">
          <Borders/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s151" ss:Parent="s121">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="Fixed"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s152" ss:Parent="s121">
          <Borders/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="Short Date"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s153" ss:Parent="s121">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="####0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s154" ss:Parent="s121">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0.0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s155" ss:Parent="s121">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s221" ss:Name="Обычный 4">
          <Alignment ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style ss:ID="s223" ss:Parent="s221">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s226" ss:Parent="s221">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s227" ss:Parent="s221">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s229" ss:Parent="s221">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s230" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s231" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s232" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="\ #############\ 0000000000000"/>
        </Style>
        <Style ss:ID="s233" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s234" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s236" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s237" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s238" ss:Parent="s221">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0000000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s239" ss:Parent="s221">
          <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s240" ss:Parent="s221">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s241" ss:Parent="s221">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="#######"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s242" ss:Parent="s221">
          <Borders/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s243" ss:Parent="s221">
          <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
          <Borders/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s244" ss:Parent="s221">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s245" ss:Parent="s221">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="000#"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s247" ss:Parent="s221">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="####\0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s248" ss:Parent="s221">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="#########\000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s249" ss:Parent="s221">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="########00000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s250" ss:Parent="s21">
          <Borders/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s251" ss:Parent="s221">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="Fixed"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s252" ss:Parent="s221">
          <Borders/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="Short Date"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s253" ss:Parent="s221">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="####0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s254" ss:Parent="s221">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0.0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s255" ss:Parent="s221">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s321" ss:Name="Обычный 5">
          <Alignment ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style ss:ID="s323" ss:Parent="s321">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s326" ss:Parent="s321">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s327" ss:Parent="s321">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s329" ss:Parent="s321">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
          </Borders>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s330" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s331" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s332" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="\ #############\ 0000000000000"/>
        </Style>
        <Style ss:ID="s333" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Color="#FF0000"
           ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s334" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s336" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss"
           ss:Color="#FF0000" ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s337" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Font ss:FontName="MS Sans Serif" x:CharSet="204" x:Family="Swiss" ss:Bold="1"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="s338" ss:Parent="s321">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0000000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s339" ss:Parent="s321">
          <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s340" ss:Parent="s321">
          <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s341" ss:Parent="s321">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="#######"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s342" ss:Parent="s321">
          <Borders/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s343" ss:Parent="s321">
          <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
          <Borders/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s344" ss:Parent="s321">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s345" ss:Parent="s321">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="000#"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s347" ss:Parent="s321">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="####\0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s348" ss:Parent="s321">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="#########\000000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s349" ss:Parent="s321">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="########00000000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s350" ss:Parent="s31">
          <Borders/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s351" ss:Parent="s321">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="Fixed"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s352" ss:Parent="s321">
          <Borders/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="Short Date"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s353" ss:Parent="s321">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="####0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s354" ss:Parent="s321">
          <Borders/>
          <Font ss:FontName="MS Sans Serif" x:Family="Swiss"/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0.0000"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s355" ss:Parent="s321">
          <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
          <Borders/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="0"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s56" ss:Parent="s21">
          <Borders/>
          <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s57" ss:Parent="s21">
          <Borders/>
          <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
        <Style ss:ID="s58" ss:Parent="s21">
          <Borders/>
          <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          <NumberFormat ss:Format="@"/>
          <Protection ss:Protected="0"/>
        </Style>
      </Styles>
      <Worksheet ss:Name="Items">
        <Table ss:ExpandedColumnCount="46" x:FullColumns="1"
         x:FullRows="1" ss:DefaultRowHeight="15">
          <Column ss:Width="79.5"/>
          <Column ss:Width="74.25"/>
          <Column ss:Hidden="1" ss:AutoFitWidth="0" ss:Span="1"/>
          <Column ss:Index="6" ss:Hidden="1" ss:AutoFitWidth="0"/>
          <Column ss:Width="128.25"/>
          <Column ss:Width="144"/>
          <Column ss:Hidden="1" ss:AutoFitWidth="0" ss:Span="1"/>
          <Column ss:Index="11" ss:Width="28.5"/>
          <Column ss:Width="72.75"/>
          <Column ss:Width="48.75"/>
          <Column ss:Width="70.5"/>
          <Column ss:Width="52.5"/>
          <Column ss:Width="41.25"/>
          <Column ss:Width="49.5"/>
          <Column ss:Width="47.25"/>
          <Column ss:Hidden="1" ss:AutoFitWidth="0" ss:Span="2"/>
          <Column ss:Index="22" ss:AutoFitWidth="0" ss:Width="144.75"/>
          <Column ss:Width="66.75"/>
          <Column ss:Width="66"/>
          <Column ss:Width="73.5"/>
          <Column ss:Width="81"/>
          <Column ss:Width="66.75"/>
          <Column ss:Hidden="1" ss:AutoFitWidth="0" ss:Span="4"/>
          <Column ss:Index="33" ss:Width="144" ss:Span="2"/>
          <Column ss:Index="36" ss:Width="101.25"/>
          <Column ss:Width="100.5"/>
          <Column ss:Width="109.5"/>
          <Column ss:Width="121.5"/>
          <Column ss:AutoFitWidth="0" ss:Width="85.5"/>
          <Column ss:AutoFitWidth="0" ss:Width="75.75"/>
          <Column ss:AutoFitWidth="0" ss:Width="55.5"/>
          <Column ss:Width="76.5"/>
          <Column ss:Width="30"/>
          <Column ss:Hidden="1" ss:AutoFitWidth="0" ss:Span="1"/>
          <Row ss:AutoFitHeight="0" ss:Height="15.75">
            <Cell ss:StyleID="s23">
              <Data ss:Type="String" x:Ticked="1">EAN ВарАрт</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String" x:Ticked="1">EAN БазВар</Data>
            </Cell>
            <Cell ss:StyleID="s26">
              <Data ss:Type="String" x:Ticked="1">НомПост</Data>
            </Cell>
            <Cell ss:StyleID="s26">
              <Data ss:Type="String" x:Ticked="1">НомПроизв</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">НомАртПост</Data>
            </Cell>
            <Cell ss:StyleID="s26">
              <Data ss:Type="String" x:Ticked="1">МиниГрТ</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String">НаимАрт I</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String">НаимАрт II</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Цвет</Data>
            </Cell>
            <Cell ss:StyleID="s26">
              <Data ss:Type="String" x:Ticked="1">ВЫГОДНО</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String" x:Ticked="1">КБУ</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String">СП</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String">ФктрСл</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String">ФктрСклОб</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Вес</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Длина</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Ширина</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Высота</Data>
            </Cell>
            <Cell ss:StyleID="s26">
              <Data ss:Type="String" x:Ticked="1">Сез</Data>
            </Cell>
            <Cell ss:StyleID="s26">
              <Data ss:Type="String" x:Ticked="1">Закупщик</Data>
            </Cell>
            <Cell ss:StyleID="s26">
              <Data ss:Type="String" x:Ticked="1">Импорт</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String">Текст на чеке</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦЗ БазВар</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦЗ ВарАрт</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦЗ действ с</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦЗ действ до</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Валюта ЦЗ</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦП БазВар</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦП ВарАрт</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦП действ с</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">ЦП действ до</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Валюта ЦП</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Дополнительный текст1</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Дополнительный текст2</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Дополнительный текст3</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String">НазвИзмБазВар</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String" x:Ticked="1">НазвИзмВарАрт</Data>
            </Cell>
            <Cell ss:StyleID="s23">
              <Data ss:Type="String" x:Ticked="1">ЕдИзмСодБазВар</Data>
            </Cell>
            <Cell ss:StyleID="s29">
              <Data ss:Type="String">Номер сертификата</Data>
            </Cell>
            <Cell ss:StyleID="s29">
              <Data ss:Type="String">Начало действия сертификата</Data>
            </Cell>
            <Cell ss:StyleID="s29">
              <Data ss:Type="String">Окончание действия сертификата</Data>
            </Cell>
            <Cell ss:StyleID="s29">
              <Data ss:Type="String">Мин. срок годности</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">Содержание</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String" x:Ticked="1">НДС</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String">Этикетка</Data>
            </Cell>
            <Cell ss:StyleID="s27">
              <Data ss:Type="String">Срок лист</Data>
            </Cell>
          </Row>
          <Row ss:AutoFitHeight="0">
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">13</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">13</Data>
            </Cell>
            <Cell ss:StyleID="s31">
              <Data ss:Type="String">8</Data>
            </Cell>
            <Cell ss:StyleID="s31">
              <Data ss:Type="String">8</Data>
            </Cell>
            <Cell ss:StyleID="s32">
              <Data ss:Type="String">14</Data>
            </Cell>
            <Cell ss:StyleID="s33">
              <Data ss:Type="String">7</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">20</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">20</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">20</Data>
            </Cell>
            <Cell ss:StyleID="s33">
              <Data ss:Type="String">1</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">4</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">3</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">4</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">4</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">9</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">4</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">4</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">4</Data>
            </Cell>
            <Cell ss:StyleID="s33">
              <Data ss:Type="String">1</Data>
            </Cell>
            <Cell ss:StyleID="s33">
              <Data ss:Type="String">8</Data>
            </Cell>
            <Cell ss:StyleID="s33">
              <Data ss:Type="String">1</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">20</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">7</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">7</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">8</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">8</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">3</Data>
            </Cell>
            <Cell ss:StyleID="s34"/>
            <Cell ss:StyleID="s34"/>
            <Cell ss:StyleID="s34"/>
            <Cell ss:StyleID="s34"/>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String">3</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String" x:Ticked="1">35</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String" x:Ticked="1">35</Data>
            </Cell>
            <Cell ss:StyleID="s34">
              <Data ss:Type="String" x:Ticked="1">35</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">3</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">3</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">3</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">20</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">8</Data>
            </Cell>
            <Cell ss:StyleID="s30">
              <Data ss:Type="String">8</Data>
            </Cell>
            <Cell ss:StyleID="s36">
              <Data ss:Type="String">4</Data>
            </Cell>
            <Cell ss:StyleID="s37">
              <Data ss:Type="String">10</Data>
            </Cell>
            <Cell ss:StyleID="s37">
              <Data ss:Type="String">2</Data>
            </Cell>
            <Cell ss:StyleID="s37">
              <Data ss:Type="String">2</Data>
            </Cell>
            <Cell ss:StyleID="s37">
              <Data ss:Type="String">8</Data>
            </Cell>
          </Row>
          <root />
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
          <PageSetup>
            <PageMargins x:Bottom="0.984251969" x:Left="0.78740157499999996"
             x:Right="0.78740157499999996" x:Top="0.984251969"/>
          </PageSetup>
          <Unsynced/>
          <Selected/>
          <Panes>
            <Pane>
              <Number>3</Number>
              <ActiveRow>11</ActiveRow>
              <ActiveCol>7</ActiveCol>
            </Pane>
          </Panes>
          <ProtectObjects>False</ProtectObjects>
          <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
      </Worksheet>
    </Workbook>
  </xsl:template>
</xsl:stylesheet>