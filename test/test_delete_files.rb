# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'

class TestDeleteFiles < Test::Unit::TestCase
  def setup
    @dir_path = 'test_delete_empty_directory'
    system("rm -rf #{@dir_path}") if FileTest.exist?(@dir_path)
  end
  
  def test_delete_empty_directory
    Dir.mkdir(@dir_path)
    assert(FileTest.exist?(@dir_path))
    Utility.delete_files(@dir_path)
    assert(!FileTest.exist?(@dir_path))
  end
  
  def test_delete_directory_and_one_file
    filename = 'test_file'
    Dir.mkdir(@dir_path)
    File.open(File.join(@dir_path, filename), "w") { |file| file.write("str") }
    assert(FileTest.exist?(@dir_path))
    assert(FileTest.exist?(File.join(@dir_path, filename)))
    Utility.delete_files(@dir_path)
    assert(!FileTest.exist?(@dir_path))
  end

  def test_delete_directory_and_subdirectory
    subdir_name = 'subdir'
    Dir.mkdir(@dir_path)
    Dir.mkdir(File.join(@dir_path, subdir_name))
    assert(FileTest.exist?(@dir_path))
    assert(FileTest.exist?(File.join(@dir_path, subdir_name)))
    Utility.delete_files(@dir_path)
    assert(!FileTest.exist?(@dir_path))
  end
end
