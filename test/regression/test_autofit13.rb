# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit13 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit13
    @xlsx = 'autofit13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_string(0, 0, 'Foo')
    worksheet.write_string(0, 1, 'Foo bar')
    worksheet.write_string(0, 2, 'Foo bar bar')

    worksheet.autofilter(0, 0, 0, 2)

    worksheet.autofit

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'],
      {}
    )
  end
end
