# -*- coding: utf-8 -*-

require 'helper'

class TestWorksheetDataBar02 < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  ###############################################################################
  #
  # Tests for WriteXLSX::Worksheet methods.
  #
  #
  def test_data_bar_02
    @worksheet.select
    @worksheet.instance_variable_set(:@index, 0)

    @worksheet.conditional_formatting(
      'A1',
      type:          'data_bar',
      data_bar_2010: 1
    )

    @worksheet.assemble_xml_file
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)

    expected = expected_to_array(expected_xml)

    assert_equal(expected, result)
  end

  def expected_xml
    <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
  <dimension ref="A1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
  <sheetData/>
  <conditionalFormatting sqref="A1">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
      <extLst>
        <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
          <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
        </ext>
      </extLst>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <extLst>
    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
      <x14:conditionalFormattings>
        <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
          <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
            <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
              <x14:cfvo type="autoMin"/>
              <x14:cfvo type="autoMax"/>
              <x14:borderColor rgb="FF638EC6"/>
              <x14:negativeFillColor rgb="FFFF0000"/>
              <x14:negativeBorderColor rgb="FFFF0000"/>
              <x14:axisColor rgb="FF000000"/>
            </x14:dataBar>
          </x14:cfRule>
          <xm:sqref>A1</xm:sqref>
        </x14:conditionalFormatting>
      </x14:conditionalFormattings>
    </ext>
  </extLst>
</worksheet>
EOS
  end
end
