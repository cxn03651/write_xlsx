# -*- coding: utf-8 -*-

require 'helper'

class TestCheckSheetname < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @workbook.add_worksheet
  end

  def test_check_sheetname_duplicate_names
    duplicate_names = %w[Sheet1 sheet1]

    duplicate_names.each do |sheetname|
      assert_raises RuntimeError, "'#{sheetname}' passed incorrectly" do
        @workbook.add_worksheet(sheetname)
      end
    end
  end

  def test_check_sheetname_invalid_characters
    invalid_characters = [
      'Sheet[', 'Sheet]', 'Sheet:', 'Sheet*', 'Sheet/', 'Sheet\\'
    ]

    invalid_characters.each do |sheetname|
      assert_raises RuntimeError, "'#{sheetname}' passed incorrectly" do
        @workbook.add_worksheet(sheetname)
      end
    end
  end

  def test_check_sheetname_long_name
    long_name = ['name_that_is_longer_than_thirty_one_characters']

    long_name.each do |sheetname|
      assert_raises RuntimeError, "'#{sheetname}' passed incorrectly" do
        @workbook.add_worksheet(sheetname)
      end
    end
  end

  def test_check_sheetname_invalid_start_stop_character
    invalid_start_end_character = ["Sheet'", "'Sheet", "'Sheet'"]

    invalid_start_end_character.each do |sheetname|
      assert_raises RuntimeError, "'#{sheetname}' passed incorrectly" do
        @workbook.add_worksheet(sheetname)
      end
    end
  end
end
