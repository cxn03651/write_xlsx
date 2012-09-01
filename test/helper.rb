# -*- coding: utf-8 -*-
require 'rubygems'
require 'bundler'
begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  $stderr.puts e.message
  $stderr.puts "Run `bundle install` to install missing gems"
  exit e.status_code
end
require 'test/unit'

$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'write_xlsx'

class Writexlsx::Workbook
  #
  # Set the default index for each format. This is mainly used for testing.
  #
  def set_default_xf_indices #:nodoc:
    @formats.each { |format| format.get_xf_index }
  end
end

class Test::Unit::TestCase
  require 'rexml/document'
  include REXML

  def setup_dir_var
    @test_dir = File.dirname(__FILE__)
    @perl_output  = File.join(@test_dir, 'perl_output')
    @regression_output  = File.join(@test_dir, 'regression', 'xlsx_files')
  end

  def expected_to_array(lines)
    array = []
    lines.each_line do |line|
      str = line.chomp.sub(%r!/>$!, ' />').sub(/^\s+/, '')
      array << str unless str == ''
    end
    array
  end

  def got_to_array(xml_str)
    str = xml_str.gsub(/[\r\n]/, '')
    str.gsub(/>[ \t\r\n]*</, ">\t<").split(/\t/)
  end

  def entrys(xlsx)
    result = []
    Zip::ZipFile.foreach(xlsx) { |entry| result << entry }
    result
  end

  def compare_xlsx(exp_filename, got_filename, ignore_members = nil, ignore_elements = nil)
    # The zip "members" are the files in the XLSX container.
    got_members = entrys(got_filename).sort_by {|member| member.name}
    exp_members = entrys(exp_filename).sort_by {|member| member.name}

    # Ignore some test specific filenames.
    if ignore_members
      got_members.reject! {|member| ignore_members.include?(member.name) }
      exp_members.reject! {|member| ignore_members.include?(member.name) }
    end

    # Check that each XLSX container has the same file members.
    assert_equal(
                 exp_members.collect {|member| member.name},
                 got_members.collect {|member| member.name},
                 "file members differs.")

    # Compare each file in the XLSX containers.
    exp_members.each_index do |i|
      got_xml_str = ""
      exp_xml_str = ""
      Document.new(got_members[i].get_input_stream.read).write(got_xml_str, 1)
      Document.new(exp_members[i].get_input_stream.read).write(exp_xml_str, 1)

      # Remove dates and user specific data from the core.xml data.
      if exp_members[i].name == 'docProps/core.xml'
        exp_xml_str = got_xml_str.gsub(/John/, '').gsub(/\d\d\d\d-\d\d-\d\dT\d\d\:\d\d:\d\dZ/,'')
        got_xml_str = exp_xml_str.gsub(/\d\d\d\d-\d\d-\d\dT\d\d\:\d\d:\d\dZ/,'')
      end

      if exp_members[i].name =~ %r!xl/worksheets/sheet\d.xml!
        exp_xml_str = exp_xml_str.sub(/horizontalDpi="200" /, '').sub(/verticalDpi="200""/, '')
        if exp_xml_str =~ /(<pageSetup.* )r:id="rId1"/
          exp_xml_str.sub(/(<pageSetup.* )r:id="rId1"/, $~[1])
        end
      end

      # Ignore test specific XML elements for defined filenames.
      if ignore_elements && ignore_elements[exp_members[i].name]
        regex = Regexp.new(ignore_elements[exp_members[i].name].
                           collect {|tag| "#{tag} [^>]+>"}.
                           join('|')
                           )
        exp_xml_str = exp_xml_str.gsub(regex, '')
        got_xml_str = got_xml_str.gsub(regex, '')
      end

      # Comparison of the XML elements in each file.
      case exp_members[i].name
      when '[Content_Types].xml', /.rels$/
        # Reorder the XML elements in the XLSX relationship files.
        assert_equal(
                     sort_rel_file_data(got_xml_str),
                     sort_rel_file_data(exp_xml_str),
                     "#{exp_members[i].name} differ."
                     )
      else
        assert_equal(
                     exp_xml_str.gsub(/ *[\r\n]+ */, ''),
                     got_xml_str.gsub(/ *[\r\n]+ */, ''),
                     "#{exp_members[i].name} differs."
                     )
      end
    end
  end

  def sort_rel_file_data(xml_str)
    array = got_to_array(xml_str.gsub(/ *[\r\n] */, ''))
    header = array.shift
    tail   = array.pop
    array.sort.unshift(header) << tail
  end
end
