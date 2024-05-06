# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx'

class TestOptionHashForWorkbook < Minitest::Test
  #
  # Workbook.new(file)
  #  => options = {}
  #  => default_format_properties = {}
  #
  def test_empty
    @workbook = WriteXLSX.new(StringIO.new)
    options = @workbook.instance_variable_get(:@options)
    default_formats = @workbook.instance_variable_get(:@default_formats)

    assert_empty(options)
    assert_empty(default_formats)
  end

  #
  # Workbook.new(file, date_1904: false)
  #  => options = {date_1904: false}
  #  => default_formats = {}
  #
  def test_one_flat_hash_with_one_key_of_options
    @workbook = WriteXLSX.new(StringIO.new, date_1904: false)
    options = @workbook.instance_variable_get(:@options)
    default_formats = @workbook.instance_variable_get(:@default_formats)

    assert_equal({ date_1904: false }, options)
    assert_empty(default_formats)
  end

  #
  # Workbook.new(file, {date_1904: false, max_url_length: 255})
  #
  #  => options = {date_1904: false}
  #  => default_formats = {color: 'red', size: 12}
  #
  def test_one_hash_with_two_keys_of_options
    @workbook = WriteXLSX.new(
      StringIO.new,
      date_1904: false, max_url_length: 255
    )
    options = @workbook.instance_variable_get(:@options)
    default_formats = @workbook.instance_variable_get(:@default_formats)

    assert_equal({ date_1904: false, max_url_length: 255 }, options)
    assert_empty(default_formats)
  end

  #
  # Workbook.new(file, date_1904: false, color: 'red' )
  #  => options = {date_1904: false}
  #  => default_formats = {color: 'red', size: 12}
  #
  def test_one_hash_with_one_option_and_one_format_option
    @workbook = WriteXLSX.new(
      StringIO.new,
      date_1904: false,
      color:     'red'
    )
    options = @workbook.instance_variable_get(:@options)
    default_formats = @workbook.instance_variable_get(:@default_formats)

    assert_equal({ date_1904: false }, options)
    assert_equal({ color: 'red' }, default_formats)
  end

  #
  # Workbook.new(file,
  #    date_1904: false, max_url_length: 255,
  #    color: 'red'
  # )
  #  => options = { date_1904: false, max_url_length: 255 }
  #  => default_formats = { color: 'red' }
  #
  def test_one_hash_with_two_options_and_one_format_option
    @workbook = WriteXLSX.new(
      StringIO.new,
      date_1904: false, max_url_length: 255,
      color: 'red'
    )
    options = @workbook.instance_variable_get(:@options)
    default_formats = @workbook.instance_variable_get(:@default_formats)

    assert_equal({ date_1904: false, max_url_length: 255 }, options)
    assert_equal({ color: 'red' }, default_formats)
  end

  #
  # Workbook.new(file,
  #    date_1904: false,
  #    color: 'red', size: 12
  # )
  #  => options = { date_1904: false }
  #  => default_formats = { color: 'red', size: 12 }
  #
  def test_one_hash_with_one_option_and_two_format_options
    @workbook = WriteXLSX.new(
      StringIO.new,
      date_1904: false,
      color: 'red', size: 12
    )
    options = @workbook.instance_variable_get(:@options)
    default_formats = @workbook.instance_variable_get(:@default_formats)

    assert_equal({ date_1904: false }, options)
    assert_equal({ color: 'red', size: 12 }, default_formats)
  end

  #
  # Workbook.new(file,
  #    date_1904: false, max_url_length: 255,
  #    color: 'red', size: 12
  # )
  #  => options = { date_1904: false, max_url_length: 255 }
  #  => default_formats = { color: 'red', size: 12 }
  #
  def test_one_hash_with_two_options_and_two_format_options
    @workbook = WriteXLSX.new(
      StringIO.new,
      date_1904: false, max_url_length: 255,
      color: 'red', size: 12
    )
    options = @workbook.instance_variable_get(:@options)
    default_formats = @workbook.instance_variable_get(:@default_formats)

    assert_equal({ date_1904: false, max_url_length: 255 }, options)
    assert_equal({ color: 'red', size: 12 }, default_formats)
  end
end
