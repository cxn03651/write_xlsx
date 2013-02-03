# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'write_xlsx/version'

Gem::Specification.new do |gem|
  gem.name          = "write_xlsx"
  gem.version       = WriteXLSX::VERSION
  gem.authors       = ["Hideo NAKAMURA"]
  gem.email         = ["cxn03651@msj.biglobe.ne.jp"]
  gem.description   = "write_xlsx s a gem to create a new file in the Excel 2007+ XLSX format, and you can use the same interface as writeexcel gem.\nThe WriteXLSX supports the following features:\n  * Multiple worksheets\n  * Strings and numbers\n  * Unicode text\n  * Rich string formats\n  * Formulas (including array formats)\n  * cell formatting\n  * Embedded images\n  * Charts\n  * Autofilters\n  * Data validation\n  * Hyperlinks\n  * Defined names\n  * Grouping/Outlines\n  * Cell comments\n  * Panes\n  * Page set-up and printing options\n\nwrite_xlsx uses the same interface as writeexcel gem.\n\ndocumentation is not completed, but writeexcel\u{2019}s documentation will help you. See http://writeexcel.web.fc2.com/\n\nAnd you can find many examples in this gem.\n"
  gem.summary       = "write_xlsx is a gem to create a new file in the Excel 2007+ XLSX format."
  gem.homepage = "http://github.com/cxn03651/write_xlsx#readme"

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ["lib"]

  gem.extra_rdoc_files = [
    "LICENSE.txt",
    "README.rdoc"
  ]

  gem.add_runtime_dependency(%q<rubyzip>, [">= 0"])
end

