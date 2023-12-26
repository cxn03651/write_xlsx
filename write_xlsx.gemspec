# frozen_string_literal: true

require_relative 'lib/write_xlsx/version'

Gem::Specification.new do |gem|
  gem.name          = 'write_xlsx'
  gem.version       = WriteXLSX_VERSION
  gem.authors       = ['Hideo NAKAMURA']
  gem.email         = ['nakamura.hideo@gmail.com']
  gem.description   = 'write_xlsx is a gem to create a new file in the Excel 2007+ XLSX format.'
  gem.summary       = 'write_xlsx is a gem to create a new file in the Excel 2007+ XLSX format.'
  gem.homepage      = 'https://github.com/cxn03651/write_xlsx#readme'
  gem.license       = 'MIT'
  gem.required_ruby_version = '>= 2.5.0'

  gem.files = Dir.chdir(File.expand_path(__dir__)) do
    `git ls-files -z`.split("\x0").reject do |f|
      (f == __FILE__) || f.match(%r{\A(?:(?:bin|test|spec|features)/|\.(?:git|travis|circleci)|appveyor)})
    end
  end
  gem.executables   = gem.files.grep(%r{^bin/}).map { |f| File.basename(f) }
  gem.require_paths = ['lib']
  gem.add_runtime_dependency 'nkf'
  gem.add_runtime_dependency 'rubyzip', '>= 1.0.0'
  gem.add_development_dependency 'byebug'
  gem.add_development_dependency 'minitest'
  gem.add_development_dependency 'mutex_m'
  gem.add_development_dependency 'rake'
  gem.add_development_dependency 'rubocop'
  gem.add_development_dependency 'rubocop-minitest'
  gem.add_development_dependency 'rubocop-rake'
  gem.extra_rdoc_files = [
    'LICENSE.txt',
    'README.md',
    'Changes'
  ]
end
