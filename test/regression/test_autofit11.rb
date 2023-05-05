# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit11 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit11
    @xlsx = 'autofit11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_formula(0, 0, "=9999+1", nil, 10000)

    worksheet.autofit

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
      {}
    )
  end
end
