# -*- coding: utf-8 -*-

require 'helper'
require 'stringio'

class TestRegressionDefinedName04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if @xlsx && File.exist?(@xlsx)
  end

  def test_defined_name_raise
    workbook = WriteXLSX.new(StringIO.new)
    worksheet = workbook.add_worksheet

    assert_raises(RuntimeError) { workbook.define_name('.abc', '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('GFG$', '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('A1',   '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('XFD1048576', '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('A A',  '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('1A',   '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('c',    '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('r',    '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('C',    '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('R',    '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('R1',   '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('C1',   '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('R1C1',   '=Sheet1!$B$1') }
    assert_raises(RuntimeError) { workbook.define_name('R13C99',   '=Sheet1!$B$1') }
  end

  def test_defined_name04
    @xlsx = 'defined_name04.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet

    # Test for valid Excel defined names.
    workbook.define_name("\\__",     '=Sheet1!$A$1')
    workbook.define_name('a3f6',     '=Sheet1!$A$2')
    workbook.define_name('afoo.bar', '=Sheet1!$A$3')
    workbook.define_name('étude',    '=Sheet1!$A$4')
    workbook.define_name('eésumé',   '=Sheet1!$A$5')
    workbook.define_name('a',        '=Sheet1!$A$6')

    workbook.close
    compare_for_regression
  end
end
