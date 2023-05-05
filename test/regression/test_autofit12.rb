# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit12 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit12
    @xlsx = 'autofit12.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_array_formula(0, 0, 2, 0, '{=SUM(B1:C1*B2:C2)}', nil, 1000)

    worksheet.write(0, 1, 20)
    worksheet.write(1, 1, 30)
    worksheet.write(2, 1, 40)

    worksheet.write(0, 2, 10)
    worksheet.write(1, 2, 40)
    worksheet.write(2, 2, 20)

    worksheet.autofit

    # Put these after the autofit() so that the autofit in on the formula result.
    worksheet.write(1, 0, 1000)
    worksheet.write(2, 0, 1000)

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'],
      {}
    )
  end
end
