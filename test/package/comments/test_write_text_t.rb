# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/comments'

class TestWriteTextT < Test::Unit::TestCase
  def setup
    @comment = Writexlsx::Package::Comments.new(nil)
  end

  def test_write_text_t_with_center_space
    @comment.__send__('write_text_t', 'Some text')
    result = @comment.xml_str
    expected = '<t>Some text</t>'
    assert_equal(expected, result)
  end

  def test_write_text_t_with_beginning_space
    @comment.__send__('write_text_t', ' Some text')
    result = @comment.xml_str
    expected = '<t xml:space="preserve"> Some text</t>'
    assert_equal(expected, result)
  end

  def test_write_text_t_with_ending_space
    @comment.__send__('write_text_t', 'Some text ')
    result = @comment.xml_str
    expected = '<t xml:space="preserve">Some text </t>'
    assert_equal(expected, result)
  end

  def test_write_text_t_with_both_space
    @comment.__send__('write_text_t', ' Some text ')
    result = @comment.xml_str
    expected = '<t xml:space="preserve"> Some text </t>'
    assert_equal(expected, result)
  end

  def test_write_text_t_with_cr
    @comment.__send__('write_text_t', "Some text\n")
    result = @comment.xml_str
    expected = %Q!<t xml:space="preserve">Some text\n</t>!
    assert_equal(expected, result)
  end
end
