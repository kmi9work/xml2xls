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
          <LastPrinted>2005-05-16T11:39:39Z</LastPrinted>
          <Created>2001-08-16T06:13:51Z</Created>
          <LastSaved>2009-10-17T23:04:43Z</LastSaved>
          <Version>12.00</Version>
        </DocumentProperties>
        <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
          <WindowHeight>9150</WindowHeight>
          <WindowWidth>15360</WindowWidth>
          <WindowTopX>0</WindowTopX>
          <WindowTopY>1365</WindowTopY>
          <TabRatio>151</TabRatio>
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
          <Style ss:ID="m38877152">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38877172">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876928">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876948">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876968">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876988">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876704">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876724">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876744">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876764">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876784">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876804">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876480">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876500">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
            <Interior/>
          </Style>
          <Style ss:ID="m38876520">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876540">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876560">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876580">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
             ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="m38876276">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Interior/>
          </Style>
          <Style ss:ID="m38876296">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Interior/>
          </Style>
          <Style ss:ID="m38876316">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Interior/>
          </Style>
          <Style ss:ID="m38876336">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Interior/>
          </Style>
          <Style ss:ID="m38876356">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Interior/>
          </Style>
          <Style ss:ID="s63">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Font ss:FontName="Arial" x:CharSet="204" x:Family="Swiss" ss:Size="14"/>
          </Style>
          <Style ss:ID="s64">
            <Borders/>
          </Style>
          <Style ss:ID="s65">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
          </Style>
          <Style ss:ID="s66">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Interior/>
          </Style>
          <Style ss:ID="s70">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Interior/>
          </Style>
          <Style ss:ID="s74">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
          </Style>
          <Style ss:ID="s75">
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="s115">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="s122">
            <Alignment ss:Vertical="Bottom" ss:Rotate="90"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
            <Interior/>
          </Style>
          <Style ss:ID="s123">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="s124">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
            <NumberFormat ss:Format="0%"/>
          </Style>
          <Style ss:ID="s125">
            <Alignment ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
            <NumberFormat ss:Format="0%"/>
          </Style>
          <Style ss:ID="s126">
            <Alignment ss:Vertical="Bottom" ss:Rotate="90" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Font ss:FontName="Arial" x:CharSet="204" ss:Size="9"/>
          </Style>
          <Style ss:ID="s127">
            <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
          </Style>
          <Style ss:ID="s128">
            <Alignment ss:Vertical="Bottom" ss:Rotate="90" ss:WrapText="1"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
          </Style>
          <Style ss:ID="s129">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
          </Style>
          <Style ss:ID="s130">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
          </Style>
          <Style ss:ID="s131">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
          </Style>
          <Style ss:ID="s132">
            <Font ss:FontName="Arial" x:Family="Swiss"/>
          </Style>
           <Style ss:ID="s133">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
           <Interior s:Color="#888888" s:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s135">
          </Style>
          <Style ss:ID="s229">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s230">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s231">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s235">
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s329">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s330">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s331">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s335">
            <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s429">
            <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s430">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s431">
            <Borders>
              <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
              <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
              <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
            </Borders>
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
          <Style ss:ID="s435">
            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
          </Style>
        </Styles>
        <Worksheet ss:Name="Форма">
          <Table ss:ExpandedColumnCount="38" x:FullColumns="1" x:FullRows="1">
            <Column ss:Width="16.5" ss:Span="2"/>
            <Column ss:Index="4" ss:Width="50.25"/>
            <Column ss:Width="174"/>
            <Column ss:Width="121.5"/>
            <Column ss:Width="28.5"/>
            <Column ss:Width="18"/>
            <Column ss:Width="12.75"/>
            <Column ss:Width="79.5" ss:Span="1"/>
            <Column ss:Index="12" ss:Width="52.5"/>
            <Column ss:Width="24" ss:Span="1"/>
            <Column ss:Index="15" ss:Width="40.5"/>
            <Column ss:Width="52.5"/>
            <Column ss:Width="18" ss:Span="1"/>
            <Column ss:Index="19" ss:Width="30" ss:Span="2"/>
            <Column ss:Index="22" ss:AutoFitWidth="0" ss:Width="34.5"/>
            <Column ss:Width="30" ss:Span="4"/>
            <Column ss:Index="28" ss:Width="84"/>
            <Column ss:Width="36.75"/>
            <Column ss:Width="27.75"/>
            <Column ss:Width="25.5"/>
            <Column ss:Width="45" ss:Span="1"/>
            <Column ss:Index="34" ss:Width="31.5" ss:Span="1"/>
            <Column ss:Index="36" ss:Width="18.75"/>
            <Column ss:AutoFitWidth="0" ss:Width="60"/>
            <Column ss:AutoFitWidth="0" ss:Width="93.75"/>
            <Row ss:AutoFitHeight="0" ss:Height="18">
              <Cell ss:MergeAcross="37" ss:StyleID="s63">
                <Data ss:Type="String">Форма предоставления информации о товаре поставщиками ЗАО &quot;Тандер&quot;</Data>
              </Cell>
            </Row>
            <Row>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:Index="11" ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
              <Cell ss:StyleID="s64"/>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="13.5"/>
            <Row ss:AutoFitHeight="0" ss:Height="13.5">
              <Cell ss:StyleID="s65">
                <Data ss:Type="Number">1</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="m38876276">
                <Data ss:Type="Number">2</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">3</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">4</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">5</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">6</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="m38876296">
                <Data ss:Type="Number">7</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">8</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">9</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">10</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="m38876316">
                <Data ss:Type="Number">11</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">12</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">13</Data>
              </Cell>
              <Cell ss:MergeAcross="10" ss:StyleID="m38876336">
                <Data ss:Type="Number">14</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">15</Data>
              </Cell>
              <Cell ss:StyleID="s66">
                <Data ss:Type="Number">16</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="m38876356">
                <Data ss:Type="Number">17</Data>
              </Cell>
              <Cell ss:MergeAcross="4" ss:StyleID="s70">
                <Data ss:Type="Number">18</Data>
              </Cell>
              <Cell ss:StyleID="s74">
                <Data ss:Type="Number">19</Data>
              </Cell>
              <Cell ss:StyleID="s74">
                <Data ss:Type="Number">20</Data>
              </Cell>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="59.25" ss:StyleID="s75">
              <Cell ss:MergeDown="1" ss:StyleID="m38876480">
                <Data ss:Type="String">Порядковый номер</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="m38876500">
                <Data ss:Type="String">Новый товар/Старый товар </Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876520">
                <Data ss:Type="String">Торговая марка на языке оригинала</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876540">
                <Data ss:Type="String">Наименование товара</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876560">
                <Data ss:Type="String">Дополнительная информация по товару (свойства, характеристики, преимущества)</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876580">
                <Data ss:Type="String">Вес, объем, емкость</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="m38876704">
                <Data ss:Type="String">Единица измерения          </Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876724">
                <Data ss:Type="String">Штрих-код товара</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876744">
                <Data ss:Type="String">Штрих-код транспортной тары (короба)</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876764">
                <Data ss:Type="String">Цена единицы товара     (с НДС)</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="m38876784">
                <Data ss:Type="String">НДС (%)</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876804">
                <Data ss:Type="String">Рекомендованная цена продажи</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876928">
                <Data ss:Type="String">Минимальная партия поставки (шт,кор., кг)</Data>
              </Cell>
              <Cell ss:MergeAcross="10" ss:StyleID="m38876948">
                <Data ss:Type="String">Габаритно-весовые параметры</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876968">
                <Data ss:Type="String">Производитель товара (на языке оригинала)</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38876988">
                <Data ss:Type="String">Страна производства</Data>
              </Cell>
              <Cell ss:MergeAcross="1" ss:StyleID="s115">
                <Data ss:Type="String">Принадлежность товара (импорт/отечественный)</Data>
              </Cell>
              <Cell ss:MergeAcross="4" ss:StyleID="s115">
                <Data ss:Type="String">Качество товара</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38877152">
                <Data ss:Type="String">Срок реализации, дней</Data>
              </Cell>
              <Cell ss:MergeDown="1" ss:StyleID="m38877172">
                <Data ss:Type="String">Валюта покупки</Data>
              </Cell>
            </Row>
            <Row ss:AutoFitHeight="0" ss:Height="126" ss:StyleID="s75">
              <Cell ss:Index="2" ss:StyleID="s122">
                <Data ss:Type="String">новинка</Data>
              </Cell>
              <Cell ss:StyleID="s122">
                <Data ss:Type="String">давно на рынке</Data>
              </Cell>
              <Cell ss:Index="8" ss:StyleID="s123">
                <Data ss:Type="String">шт.</Data>
              </Cell>
              <Cell ss:StyleID="s123">
                <Data ss:Type="String">кг</Data>
              </Cell>
              <Cell ss:Index="13" ss:StyleID="s124">
                <Data ss:Type="Number">0.1</Data>
              </Cell>
              <Cell ss:StyleID="s125">
                <Data ss:Type="Number">0.18</Data>
              </Cell>
              <Cell ss:Index="17" ss:StyleID="s126">
                <Data ss:Type="String">шт. в блоке/спайке</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">шт. в коробке</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">коробов в слое</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">коробов на EURO-паллете</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">габариты блока/спайки (дл./шир./высота)</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">габариты короба (дл./шир./высота)</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">высота EURO-паллеты</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">вес брутто ед. товара</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">вес нетто ед. товара</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">вес брутто короба (кг)</Data>
              </Cell>
              <Cell ss:StyleID="s126">
                <Data ss:Type="String">вес брутто поддона (кг)</Data>
              </Cell>
              <Cell ss:Index="30" ss:StyleID="s127">
                <Data ss:Type="String">отеч.</Data>
              </Cell>
              <Cell ss:StyleID="s127">
                <Data ss:Type="String">имп.</Data>
              </Cell>
              <Cell ss:StyleID="s128">
                <Data ss:Type="String">код ОКПД/ТНВЭД (сертификат соответствия)</Data>
              </Cell>
              <Cell ss:StyleID="s128">
                <Data ss:Type="String">порядковый номер сертификата соответствия</Data>
              </Cell>
              <Cell ss:StyleID="s128">
                <Data ss:Type="String">гигиенический сертификат</Data>
              </Cell>
              <Cell ss:StyleID="s128">
                <Data ss:Type="String">удостоверение качества</Data>
              </Cell>
              <Cell ss:StyleID="s128">
                <Data ss:Type="String">ГТД</Data>
              </Cell>
            </Row>
            <root />
          </Table>
          <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
            <PageSetup>
              <Layout x:Orientation="Landscape"/>
              <Header x:Margin="0.59055118110236227"/>
              <Footer x:Margin="0.51181102362204722"/>
              <PageMargins x:Bottom="0.39370078740157483" x:Left="0.19685039370078741"
               x:Right="0.19685039370078741" x:Top="0.39370078740157483"/>
            </PageSetup>
            <FitToPage/>
            <Print>
              <ValidPrinterInfo/>
              <PaperSizeIndex>9</PaperSizeIndex>
              <Scale>60</Scale>
              <HorizontalResolution>1200</HorizontalResolution>
              <VerticalResolution>1200</VerticalResolution>
            </Print>
            <PageBreakZoom>60</PageBreakZoom>
            <Selected/>
            <TopRowVisible>3</TopRowVisible>
            <LeftColumnVisible>1</LeftColumnVisible>
            <Panes>
              <Pane>
                <Number>3</Number>
                <ActiveRow>5</ActiveRow>
                <ActiveCol>30</ActiveCol>
                <RangeSelection>C31</RangeSelection>
              </Pane>
            </Panes>
            <ProtectObjects>False</ProtectObjects>
            <ProtectScenarios>False</ProtectScenarios>
          </WorksheetOptions>
        </Worksheet>
      </Workbook>
    </xsl:template>
</xsl:stylesheet>
