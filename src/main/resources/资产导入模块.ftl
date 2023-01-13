<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Apache POI</Author>
  <LastAuthor>shidang deng</LastAuthor>
  <Created>2023-01-12T06:06:40Z</Created>
  <LastSaved>2023-01-12T11:32:38Z</LastSaved>
  <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>8580</WindowHeight>
  <WindowWidth>21570</WindowWidth>
  <WindowTopX>32767</WindowTopX>
  <WindowTopY>32767</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
                                   <Borders/>
                                            <Font ss:FontName="宋体" x:CharSet="134" ss:Size="11" ss:Color="#000000"/>
                                                                                                                   <Interior/>
                                                                                                                             <NumberFormat/>
                                                                                                                                           <Protection/>
  </Style>
  <Style ss:ID="s63">
   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
  </Style>
  <Style ss:ID="s64">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
                                                          <Borders>
                                                          <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      </Borders>
                        <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FF0000"/>
                                                                                                     <Interior ss:Color="#00CCFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
                                                          <Borders>
                                                          <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      </Borders>
                        <Font ss:FontName="宋体" x:CharSet="134" ss:Size="11" ss:Color="#FF0000"/>
                                                                                               <Interior ss:Color="#00CCFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s66">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
                                                          <Borders>
                                                          <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
   ss:Color="#000000"/>
                      </Borders>
                        <Interior ss:Color="#00CCFF" ss:Pattern="Solid"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="资产导入模板">
  <Table ss:ExpandedColumnCount="22" ss:ExpandedRowCount="2005" x:FullColumns="1"
         x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="13.5">
   <Column ss:AutoFitWidth="0" ss:Width="157.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="162"/>
   <Row ss:AutoFitHeight="0" ss:Height="90">
    <Cell ss:MergeAcross="21" ss:StyleID="s63"><Data ss:Type="String">填写说明&#13;&#10;1、红色字体为必填项，黑色字体为选填项&#13;&#10;2、多级用/隔开，例如：制造事业部/北京分部/设计部&#13;&#10;3、您已开启自动生成资产编码，资产编码字段无需输入&#13;&#10;</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="15">
    <Cell ss:StyleID="s64"><Data ss:Type="String">资产分类</Data><Comment ss:Author=""><ss:Data
              xmlns="http://www.w3.org/TR/REC-html40"><Font html:Size="9"
                                                            html:Color="#000000">必填</Font></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">资产名称</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">品牌</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">型号</Data><Comment ss:Author=""><ss:Data
              xmlns="http://www.w3.org/TR/REC-html40"><Font html:Size="9"
                                                            html:Color="#000000">必填</Font></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">设备序列号</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">管理员</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">所属公司</Data><Comment ss:Author=""><ss:Data
              xmlns="http://www.w3.org/TR/REC-html40"><Font html:Size="9"
                                                            html:Color="#000000">必填</Font></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">所在位置</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">购置时间</Data><Comment ss:Author=""><ss:Data
              xmlns="http://www.w3.org/TR/REC-html40"><Font html:Size="9"
                                                            html:Color="#000000">必填</Font></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">购置方式</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">购置金额（含税）</Data><Comment
             ss:Author=""><ss:Data xmlns="http://www.w3.org/TR/REC-html40"><Font
               html:Size="9" html:Color="#000000">必填</Font></ss:Data></Comment></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">入库时间</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">预计使用期限(月)</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">备注</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">资产照片</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">使用人</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">使用部门</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">领用日期</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">保养到期时间</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">保养说明</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">预计折旧期限(月)</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">入库数量</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <#--    <Cell><Data ss:Type="String">计算机和网络设备/笔记本电脑</Data></Cell>-->
    <#--    <Cell><Data ss:Type="String">键盘</Data></Cell>-->
    <#--    <Cell><Data ss:Type="String">罗技</Data></Cell>-->
    <#--    <Cell ss:Index="6"><Data ss:Type="String">York</Data></Cell>-->
    <#--    <Cell><Data ss:Type="String">去玩吧</Data></Cell>-->
    <#--    <Cell><Data ss:Type="String">北京</Data></Cell>-->
    <#--    <Cell ss:Index="10"><Data ss:Type="String">租赁</Data></Cell>-->
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>13</ActiveRow>
     <ActiveCol>6</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
   <Range>R3C17:R20001C17</Range>
   <Type>List</Type>
   <CellRangeList/>
   <#--   <Value>去玩吧,去玩吧/销售部,去玩吧/人事部,去玩吧/采购部,去玩吧/研发部</Value>-->
   <Value>${detpNames}</Value>
   <InputHide/>
  </DataValidation>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
   <Range>R3C6:R20001C6</Range>
   <Type>List</Type>
   <CellRangeList/>
   <#--   <Value>York</Value>-->
   <Value>${admins}</Value>
   <InputHide/>
  </DataValidation>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
   <Range>R3C7:R20001C7</Range>
   <Type>List</Type>
   <CellRangeList/>
   <#--   <Value>去玩吧</Value>-->
   <Value>${companys}</Value>
   <InputHide/>
  </DataValidation>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
   <Range>R3C8:R20001C8</Range>
   <Type>List</Type>
   <CellRangeList/>
   <#--   <Value>北京,深圳</Value>-->
   <Value>${godownNames}</Value>
   <InputHide/>
  </DataValidation>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
   <Range>R3C10:R20001C10</Range>
   <Type>List</Type>
   <CellRangeList/>
   <#--   <Value>采购,租赁</Value>-->
   <#--   <Value>${assetsSources}</Value>-->
   <Value>${assetsSources}</Value>

   <InputHide/>
  </DataValidation>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
   <Range>R3C1:R20001C1</Range>
   <Type>List</Type>
   <CellRangeList/>
   <#--   <Value>办公设备/LED显示屏,办公设备/标签机,办公设备/会计器具,办公设备/其他办公设备,办公家具/桌子,办公家具/椅子,办公家具/沙发,办公家具/文件柜,办公家具/保险柜,办公家具/其他柜</Value>-->
   <#--   <Value>${assetsTypeNames}</Value>-->
   <Value>${assetsTypeNames}</Value>
  </DataValidation>
 </Worksheet>
</Workbook>
