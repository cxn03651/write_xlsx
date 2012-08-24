# encoding: utf-8

require 'rubygems'
require 'bundler'
begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  $stderr.puts e.message
  $stderr.puts "Run `bundle install` to install missing gems"
  exit e.status_code
end
require 'rake'

require 'jeweler'
Jeweler::Tasks.new do |gem|
  # gem is a Gem::Specification... see http://docs.rubygems.org/read/chapter/20 for more options
  gem.name = "write_xlsx"
  gem.homepage = "http://github.com/cxn03651/write_xlsx"
  gem.license = "MIT"
  gem.summary = %Q{write_xlsx is a gem to create a new file in the Excel 2007+ XLSX format.}
  gem.description = <<EOS
write_xlsx s a gem to create a new file in the Excel 2007+ XLSX format, and you can use the same interface as writeexcel gem.
The WriteXLSX supports the following features:
  * Multiple worksheets
  * Strings and numbers
  * Unicode text
  * Rich string formats
  * Formulas (including array formats)
  * cell formatting
  * Embedded images
  * Charts
  * Autofilters
  * Data validation
  * Hyperlinks
  * Defined names
  * Grouping/Outlines
  * Cell comments
  * Panes
  * Page set-up and printing options

write_xlsx uses the same interface as writeexcel gem.

documentation is not completed, but writeexcel’s documentation will help you. See http://writeexcel.web.fc2.com/

And you can find many examples in this gem.
EOS
  gem.email = "cxn03651@msj.biglobe.ne.jp"
  gem.authors = ["Hideo NAKAMURA"]
  # dependencies defined in Gemfile
end
Jeweler::RubygemsDotOrgTasks.new

require 'rake/testtask'
Rake::TestTask.new(:test) do |test|
  test.libs << 'lib' << 'test'
  test.pattern = 'test/**/test_*.rb'
  test.verbose = true
end

# require 'rcov/rcovtask'
# Rcov::RcovTask.new do |test|
#   test.libs << 'test'
#   test.pattern = 'test/**/test_*.rb'
#   test.verbose = true
#   test.rcov_opts << '--exclude "gems/*"'
# end

task :default => :test

require 'rake/rdoctask'
Rake::RDocTask.new do |rdoc|
  version = File.exist?('VERSION') ? File.read('VERSION') : ""

  rdoc.rdoc_dir = 'rdoc'
  rdoc.title = "write_xlsx #{version}"
  rdoc.rdoc_files.include('README*')
  rdoc.rdoc_files.include('lib/**/*.rb')
end
