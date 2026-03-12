# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx'

class TestDeleteTempdir < Minitest::Test
  def setup
    @dir_path = 'test_delete_empty_directory'
    system("rm -rf #{@dir_path}") if FileTest.exist?(@dir_path)
  end

  def test_delete_tempdir_removes_empty_directory
    Dir.mkdir(@dir_path)

    assert(FileTest.exist?(@dir_path))

    workbook = Writexlsx::Workbook.new(StringIO.new)
    workbook.send(:delete_tempdir, @dir_path)

    refute(FileTest.exist?(@dir_path))
  end

  def test_delete_tempdir_removes_directory_and_one_file
    filename = 'test_file'
    Dir.mkdir(@dir_path)
    File.write(File.join(@dir_path, filename), "str")

    assert(FileTest.exist?(@dir_path))
    assert(FileTest.exist?(File.join(@dir_path, filename)))

    workbook = Writexlsx::Workbook.new(StringIO.new)
    workbook.send(:delete_tempdir, @dir_path)

    refute(FileTest.exist?(@dir_path))
  end

  def test_delete_tempdir_removes_directory_and_subdirectory
    subdir_name = 'subdir'
    Dir.mkdir(@dir_path)
    Dir.mkdir(File.join(@dir_path, subdir_name))

    assert(FileTest.exist?(@dir_path))
    assert(FileTest.exist?(File.join(@dir_path, subdir_name)))

    workbook = Writexlsx::Workbook.new(StringIO.new)
    workbook.send(:delete_tempdir, @dir_path)

    refute(FileTest.exist?(@dir_path))
  end
end
