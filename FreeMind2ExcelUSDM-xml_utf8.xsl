<?xml version="1.0" encoding="UTF-8"?>
<!--
    (c) 2018 Shoichi Hayashi(林 祥一)
    このコードはGPLv3の元にライセンスします。
    (http://www.gnu.org/copyleft/gpl.html)
    FreeMindで一定のルールに従って記述したUSDM流の要求仕様書を
    ExcelのUSDMとして読み込み可能なXMLファイルに変換します。
    （以下Windowsでの使い方を示します。）
    FreeMindの以下のメニューから使用します。
    　ファイル -> エクスポート -> XSLTを使用...
    出てきたダイアログの「XSLファイルを選択」でこのファイルを指定します。
    「エクスポートファイルを選択」の方に出力先ファイルを指定します。
    （拡張子は.xmlが良いでしょう。）そして「エクスポート」ボタンを押します。
    変換出力されたxmlファイルを右クリックし、出てきたメニューから
    　プログラムから開く -> Excel
    以上で、Excelが起動してUSDMが読み込まれます。
-->

<xsl:stylesheet version="1.0"
 xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
 xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <xsl:output method="xml" indent="yes" encoding="UTF-8" standalone="yes" cdata-section-elements="Data" />
  <xsl:variable name="UsdmMaxLevel" select="8"/>
  <xsl:variable name="BodyIcon" select="'yes'"/>
  <xsl:variable name="BodyColor" select="'#EEEEEE'"/>
  <xsl:variable name="UsdmTitle" select="'USDMによる要求仕様'"/>
  <xsl:variable name="TitleFontColor" select="'#FFFFFF'"/>
  <xsl:variable name="TitleColor" select="'#000000'"/>
  <xsl:variable name="RequirementTitle" select="'要求'"/>
  <xsl:variable name="RequirementIcon" select="'flag-green'"/>
  <xsl:variable name="RequirementColor" select="'#C1FFC1'"/>
  <xsl:variable name="ReasonTitle" select="'理由'"/>
  <xsl:variable name="ReasonIcon" select="'help'"/>
  <xsl:variable name="ExplanationTitle" select="'説明'"/>
  <xsl:variable name="ExplanationIcon" select="'info'"/>
  <xsl:variable name="SpecificationTitle" select="'□□□□□'"/>
  <xsl:variable name="SpecificationIcon" select="'flag-blue'"/>
  <xsl:variable name="SpecificationColor" select="'#B3EBFF'"/>
  <xsl:variable name="GroupIcon" select="'folder'"/>
  <xsl:variable name="GroupColor" select="'#CECEFF'"/>
  <xsl:variable name="RemarkTitle" select="'備考'"/>
  <xsl:variable name="RemarkIcon" select="'pencil'"/>
  <xsl:variable name="RemarkWidth" select="240"/>
  <xsl:variable name="DefaultWidth" select="80"/>

  <xsl:template match="/map">
    <xsl:variable name="SheetName" select="node/@TEXT[1]"/>
    <xsl:processing-instruction name="mso-application"> progid="Excel.Sheet"</xsl:processing-instruction>
    <Workbook>
      <Styles>
        <Style ss:ID="sTitle">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center" />
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Font ss:Color="{$TitleFontColor}" ss:Bold="1"/>
          <Interior ss:Color="{$TitleColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sRequirement">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$RequirementColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sReqId">
          <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
          <Borders>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$RequirementColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sReqBody">
          <Alignment ss:Vertical="Top" ss:WrapText="1"/>
          <Borders>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$BodyColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sReason">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous"/>
          </Borders>
          <Interior ss:Color="{$RequirementColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sRsnBody">
          <Alignment ss:Vertical="Top" ss:WrapText="1"/>
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous"/>
          </Borders>
          <Interior ss:Color="{$BodyColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sExplanation">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$RequirementColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sExplnBody">
          <Alignment ss:Vertical="Top" ss:WrapText="1"/>
          <Borders>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$BodyColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sGroup">
          <Alignment ss:Vertical="Center"/>
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$GroupColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sSpecification">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center" />
        </Style>
        <Style ss:ID="sSpecId">
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$SpecificationColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sBody">
          <Alignment ss:Vertical="Top" ss:WrapText="1"/>
          <Borders>
            <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
            <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
          </Borders>
          <Interior ss:Color="{$BodyColor}" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="sRemarkBody">
          <Alignment ss:Vertical="Top" ss:WrapText="1"/>
          <Interior ss:Color="{$BodyColor}" ss:Pattern="Solid"/>
        </Style>
      </Styles>
      <Worksheet ss:Name="{$SheetName}">
        <Table            ss:DefaultColumnWidth="{$DefaultWidth}">
          <Column ss:Index="{$UsdmMaxLevel+3}" ss:Width="{$RemarkWidth}"/>
          <Row>
            <Cell ss:StyleID="sTitle" ss:Index="1" ss:MergeAcross="{$UsdmMaxLevel+1}">
              <Data ss:Type="String">
                <xsl:value-of select="$UsdmTitle"/>
              </Data>
            </Cell>
            <Cell ss:StyleID="sTitle" ss:Index="{$UsdmMaxLevel+3}">
              <Data ss:Type="String">
                <xsl:value-of select="$RemarkTitle"/>
              </Data>
            </Cell>
          </Row>
          <xsl:apply-templates select="node/node">
            <xsl:with-param name="level" select="1" />
          </xsl:apply-templates>
        </Table>
      </Worksheet>
    </Workbook>
  </xsl:template>

  <xsl:template match="node">
    <xsl:param name="level" />
    <xsl:param name="id" />
    <xsl:param name="kind" />
    <xsl:choose>
      <xsl:when test="icon[@BUILTIN=$RequirementIcon]">
        <xsl:apply-templates select="node">
          <xsl:with-param name="level" select="$level + 1" />
          <xsl:with-param name="id">
            <xsl:call-template name="textData" />
          </xsl:with-param>
          <xsl:with-param name="kind" select="$RequirementTitle"/>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:when test="icon[@BUILTIN=$SpecificationIcon]">
        <xsl:apply-templates select="node">
          <xsl:with-param name="level" select="$level" />
          <xsl:with-param name="id">
            <xsl:call-template name="textData" />
          </xsl:with-param>
          <xsl:with-param name="kind" select="$SpecificationTitle"/>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:when test="icon[@BUILTIN=$BodyIcon]">
        <Row>
          <xsl:choose>
            <xsl:when test="$kind=$RequirementTitle">
              <Cell ss:StyleID="sRequirement" ss:Index="{$level - 1}">
                <Data ss:Type="String">
                  <xsl:value-of select="$kind"/>
                </Data>
              </Cell>
              <Cell ss:StyleID="sReqId" ss:Index="{$level}">
                <Data ss:Type="String">
                  <xsl:value-of select="$id"/>
                </Data>
              </Cell>
              <Cell ss:StyleID="sReqBody" ss:Index="{$level + 1}" ss:MergeAcross="{$UsdmMaxLevel+1 - $level}">
                <xsl:call-template name="textData" />
              </Cell>
            </xsl:when>
            <xsl:when test="$kind=$SpecificationTitle">
              <Cell ss:StyleID="sSpecification" ss:Index="{$level - 1}">
                <Data ss:Type="String">
                  <xsl:value-of select="$kind"/>
                </Data>
              </Cell>
              <Cell ss:StyleID="sSpecId" ss:Index="{$level}">
                <Data ss:Type="String">
                  <xsl:value-of select="$id"/>
                </Data>
              </Cell>
              <Cell ss:StyleID="sBody" ss:Index="{$level + 1}" ss:MergeAcross="{$UsdmMaxLevel+1 - $level}">
                <xsl:call-template name="textData" />
              </Cell>
            </xsl:when>
          </xsl:choose>
          <xsl:apply-templates select="node">
            <xsl:with-param name="level" select="$level" />
            <xsl:with-param name="id">
              leaf
            </xsl:with-param>
          </xsl:apply-templates>
        </Row>
        <xsl:apply-templates select="node">
          <xsl:with-param name="level" select="$level" />
          <xsl:with-param name="id">
            foobar
          </xsl:with-param>
          <xsl:with-param name="kind" select="$kind" />
        </xsl:apply-templates>
      </xsl:when>
      <xsl:when test="icon[@BUILTIN=$ReasonIcon]">
         <Row>
          <Cell ss:StyleID="sReason" ss:Index="{$level}">
            <Data ss:Type="String">
              <xsl:value-of select="$ReasonTitle"/>
            </Data>
          </Cell>
          <Cell ss:StyleID="sRsnBody" ss:Index="{$level + 1}" ss:MergeAcross="{$UsdmMaxLevel+1 - $level}">
            <xsl:call-template name="textData" />
          </Cell>
          <xsl:apply-templates select="node">
            <xsl:with-param name="level" select="$level" />
            <xsl:with-param name="id">
              leaf
            </xsl:with-param>
            <xsl:with-param name="kind" select="$kind" />
          </xsl:apply-templates>
        </Row>
      </xsl:when>
      <xsl:when test="icon[@BUILTIN=$ExplanationIcon]">
         <Row>
          <Cell ss:StyleID="sExplanation" ss:Index="{$level}">
            <Data ss:Type="String">
              <xsl:value-of select="$ExplanationTitle"/>
            </Data>
          </Cell>
          <Cell ss:StyleID="sExplnBody" ss:Index="{$level + 1}" ss:MergeAcross="{$UsdmMaxLevel+1 - $level}">
            <xsl:call-template name="textData" />
          </Cell>
          <xsl:apply-templates select="node">
            <xsl:with-param name="level" select="$level" />
            <xsl:with-param name="id">
              leaf
            </xsl:with-param>
            <xsl:with-param name="kind" select="$kind" />
          </xsl:apply-templates>
        </Row>
      </xsl:when>
      <xsl:when test="icon[@BUILTIN=$GroupIcon]">
        <Row>
          <Cell ss:StyleID="sGroup" ss:Index="{$level}" ss:MergeAcross="{$UsdmMaxLevel+1 + 1 - $level}">
            <xsl:call-template name="textData" />
          </Cell>
        </Row>
        <xsl:apply-templates select="node">
          <xsl:with-param name="level" select="$level" />
          <xsl:with-param name="id">
            group
          </xsl:with-param>
          <xsl:with-param name="kind" select="$kind" />
        </xsl:apply-templates>
      </xsl:when>
      <xsl:when test="icon[@BUILTIN=$RemarkIcon]">
        <xsl:if test="contains($id, 'leaf')">
          <Cell ss:StyleID="sRemarkBody" ss:Index="{$UsdmMaxLevel+3}">
            <xsl:call-template name="textData" />
          </Cell>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="textData">
    <xsl:choose>
      <xsl:when test="richcontent[@TYPE='NODE']">
        <Data ss:Type="String">
          <xsl:copy-of select="richcontent[@TYPE='NODE']/html/body/*" />
        </Data>
      </xsl:when>
      <xsl:otherwise>
        <Data ss:Type="String">
        <xsl:value-of select="@TEXT"/>
    </Data>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>