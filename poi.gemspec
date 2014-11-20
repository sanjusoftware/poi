# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'poi/version'

Gem::Specification.new do |spec|
  spec.name          = "poi"
  spec.version       = Poi::VERSION
  spec.authors       = ['Sanjeev Mishra']
  spec.email         = ['sanjusoftware@gmail.com']
  spec.summary       = %q{A Ruby gem to work with MS Office suit using apache poi}
  spec.description   = %q{poi is a Ruby gem that can help working with MS Office suit using the apache poi.}
  spec.homepage      = ""
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency 'bundler', "~> 1.5"
  spec.add_development_dependency 'rake'

  spec.add_dependency 'rjb'
end
