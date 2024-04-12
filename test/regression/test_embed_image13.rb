# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# Conver to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'helper'

class TestRegressionEmbedImage13 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_embed_image13
    @xlsx = 'embed_image13.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet

    worksheet1.embed_image(0, 0, 'test/regression/images/red.png')
    worksheet1.embed_image(2, 0, 'test/regression/images/blue.png')
    worksheet1.embed_image(4, 0, 'test/regression/images/yellow.png')

    worksheet2 = workbook.add_worksheet

    worksheet2.embed_image(0, 0, 'test/regression/images/yellow.png')
    worksheet2.embed_image(2, 0, 'test/regression/images/red.png')
    worksheet2.embed_image(4, 0, 'test/regression/images/blue.png')

    worksheet3 = workbook.add_worksheet

    worksheet3.embed_image(0, 0, 'test/regression/images/blue.png')
    worksheet3.embed_image(2, 0, 'test/regression/images/yellow.png')
    worksheet3.embed_image(4, 0, 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
