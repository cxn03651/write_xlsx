# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionShapeScale01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_shape_scale01
    @xlsx = 'shape_scale01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    format      = workbook.add_format(:font => 'Arial', :size => 8)

    normal      = workbook.add_shape(
                                     :name   => 'chip',
                                     :type   => 'diamond',
                                     :text   => 'Normal',
                                     :width  => 100,
                                     :height => 100,
                                     :format => format
                                     )
    worksheet.insert_shape('A1', normal, 50, 50)

    normal.text = 'Scaled 3w x 2h'
    normal.name = 'Hope'
    worksheet.insert_shape('A1', normal, 250, 50, 3, 2)

    workbook.close
    compare_for_regression(
      %w[
        xl/printerSettings/printerSettings1.bin
        xl/worksheets/_rels/sheet1.xml.rels
      ],
      {
        '[Content_Types].xml'      => ['<Default Extension="bin"'],
        'xl/worksheets/sheet1.xml' => ['<pageMargins']
      }
    )
  end
end
