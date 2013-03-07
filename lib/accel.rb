require 'nokogiri'
require 'zip/zip'

project_root  = File.join(File.expand_path(File.dirname(__FILE__)), '..')
examples_root = File.join(project_root, 'examples')
tmp_root = File.join(project_root, 'tmp')

class FakeData
  def self.title
    'hello world'
  end
end

class TemplateSpreadsheet
  attr_reader :xml

  def initialize(filepath)
    @xml = Nokogiri::XML(File.read(filepath))
  end

  def cells
    @xml.css("Cell Data")
  end

  def dynamic_cells
    cells.select { |cell|
      cell.content.match /^{{(.*?)}}$/
    }
  end

  def substitute_contents
    dynamic_cells.each do |cell|
      key = cell.content.gsub('{{','').gsub('}}','')
      cell.content = FakeData.send(key.to_sym)
    end
  end

  def write_tempfile(filepath)
    tempfile = File.open(filepath, 'w') do |f|
      f.write @xml.to_s
    end
  end
end

# Load it
filepath = File.join(examples_root, "attribute.xml")
book = TemplateSpreadsheet.new(filepath)
book.substitute_contents
book.write_tempfile(File.join(tmp_root, 'attribute_new.xml'))
