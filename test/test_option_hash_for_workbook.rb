# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx'

class ForTest
  include Writexlsx::Utility
end

class TestOptionHashForWorkbook < Minitest::Test
  def setup
    @obj = ForTest.new
    @options = {
      :tempdir         => 'temp',
      :date_1904       => false,
      :optimization    => false,
      :excel2003_style => false
    }
    @default_format_properties = { :size => 12, :color => 'red' }
  end

  #
  # Workbook.new(file)
  #  => options = {}
  #  => default_format_properties = {}
  #
  def test_empty
    options, default_format_properties = @obj.process_workbook_options

    assert_empty(options)
    assert_empty(default_format_properties)
  end

  #
  # Workbook.new(file, date_1904: false, color: 'red', size: 12)
  #  => options = {date_1904: false}
  #  => default_formats = {color: 'red', size: 12}
  #
  def test_one_flat_hash
    params = @options.merge(@default_format_properties)
    options, default_format_properties = @obj.process_workbook_options(params)

    assert_equal(@options, options)
    assert_equal(@default_format_properties, default_format_properties)
  end

  #
  # Workbook.new(file, date_1904: false,
  #              default_format_properties: {color: 'red', size: 12} )
  #  => options = {date_1904: false}
  #  => default_formats = {color: 'red', size: 12}
  #
  def test_one_hash_includes_format_key
    params = @options.dup
    params[:default_format_properties] = @default_format_properties

    options, default_format_properties = @obj.process_workbook_options(params)

    assert_equal(@options, options)
    assert_equal(@default_format_properties, default_format_properties)
  end

  #
  # Workbook.new(file, {date_1904: false},
  #                    {color: 'red', size: 12} )
  #  => options = {date_1904: false}
  #  => default_formats = {color: 'red', size: 12}
  #
  def test_two_hash
    options, default_format_properties =
      @obj.process_workbook_options(@options, @default_format_properties)

    assert_equal(@options, options)
    assert_equal(@default_format_properties, default_format_properties)
  end
end
