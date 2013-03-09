require 'nokogiri'
require 'zip/zip'
require_relative 'template_spreadsheet'

project_root  = File.join(File.expand_path(File.dirname(__FILE__)), '..')
examples_root = File.join(project_root, 'examples')
tmp_root = File.join(project_root, 'tmp')

class FakeData
  def self.title
    'hello world'
  end
end

# Load it
filepath = File.join(examples_root, "rows.xml")
book = TemplateSpreadsheet.new(filepath)
puts book.ast.map(&:inspect)
# book.substitute_contents
# book.write_tempfile(File.join(tmp_root, 'attribute_new.xml'))
