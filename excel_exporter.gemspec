# Provide a simple gemspec so you can easily use your enginex
# project in your rails apps through git.
Gem::Specification.new do |s|
  s.name = "excel_exporter"
  s.summary = "Insert ExcelExporter summary."
  s.description = "Insert ExcelExporter description."
  s.files = Dir["{app,lib,config}/**/*"] + ["MIT-LICENSE", "Rakefile", "Gemfile", "README.rdoc"]
  s.version = "0.0.2"
  s.author = 'Gregory Shilin'
  s.add_dependency('htmlentities')
end