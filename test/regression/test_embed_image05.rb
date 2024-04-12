# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# Conver to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'helper'

class TestRegressionEmbedImage05 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_embed_image05
    @xlsx = 'embed_image05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_dynamic_array_formula(0, 0, 2, 0, '=LEN(B1:B3)', nil, 0)

    worksheet.embed_image(8, 4, 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
