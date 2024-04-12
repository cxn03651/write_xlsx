# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# Conver to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'helper'

class TestRegressionEmbedImage08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_embed_image08
    @xlsx = 'embed_image08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.embed_image(
      0, 0, 'test/regression/images/red.png',
      description: "Some alt text"
    )

    workbook.close
    compare_for_regression
  end
end
