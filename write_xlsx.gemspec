# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'write_xlsx/version'

Gem::Specification.new do |gem|
  gem.name          = "write_xlsx"
  gem.version       = WriteXLSX::VERSION
  gem.authors       = ["Hideo NAKAMURA"]
  gem.email         = ["cxn03651@msj.biglobe.ne.jp"]
  gem.description   = "write_xlsx is a gem to create a new file in the Excel 2007+ XLSX format."
  gem.summary       = "write_xlsx is a gem to create a new file in the Excel 2007+ XLSX format."
  gem.homepage = "http://github.com/cxn03651/write_xlsx#readme"

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ["lib"]

  gem.extra_rdoc_files = [
    "LICENSE.txt",
    "README.md",
    "Changes"
  ]

  gem.add_runtime_dependency(%q<rubyzip>, [">=1.0.0"])
  gem.add_development_dependency(%q<simplecov>) if RUBY_VERSION >= "1.9"
end
