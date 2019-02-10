# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionPanes01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_panes01
    @xlsx = 'panes01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet01 = workbook.add_worksheet
    worksheet02 = workbook.add_worksheet
    worksheet03 = workbook.add_worksheet
    worksheet04 = workbook.add_worksheet
    worksheet05 = workbook.add_worksheet
    worksheet06 = workbook.add_worksheet
    worksheet07 = workbook.add_worksheet
    worksheet08 = workbook.add_worksheet
    worksheet09 = workbook.add_worksheet
    worksheet10 = workbook.add_worksheet
    worksheet11 = workbook.add_worksheet
    worksheet12 = workbook.add_worksheet
    worksheet13 = workbook.add_worksheet

    worksheet01.write('A1', 'Foo')
    worksheet02.write('A1', 'Foo')
    worksheet03.write('A1', 'Foo')
    worksheet04.write('A1', 'Foo')
    worksheet05.write('A1', 'Foo')
    worksheet06.write('A1', 'Foo')
    worksheet07.write('A1', 'Foo')
    worksheet08.write('A1', 'Foo')
    worksheet09.write('A1', 'Foo')
    worksheet10.write('A1', 'Foo')
    worksheet11.write('A1', 'Foo')
    worksheet12.write('A1', 'Foo')
    worksheet13.write('A1', 'Foo')

    worksheet01.freeze_panes('A2')
    worksheet02.freeze_panes('A3')
    worksheet03.freeze_panes('B1')
    worksheet04.freeze_panes('C1')
    worksheet05.freeze_panes('B2')
    worksheet06.freeze_panes('G4')
    worksheet07.freeze_panes(3, 6, 3, 6, 1)
    worksheet08.split_panes(15, 0)
    worksheet09.split_panes(30, 0)
    worksheet10.split_panes(0, 8.46)
    worksheet11.split_panes(0, 17.57)
    worksheet12.split_panes(15, 8.46)
    worksheet13.split_panes(45, 54.14)

    workbook.close
    compare_for_regression
  end
end
