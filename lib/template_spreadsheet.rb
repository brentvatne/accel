class TemplateSpreadsheet
  attr_reader :xml

  def initialize(filepath)
    @xml = Nokogiri::XML(File.read(filepath))
  end

  def cell_nodes
    @xml.css("Cell Data")
  end

  def dynamic_cells
    cell_nodes.select { |node|
      node.content.match /{{(.*?)}}/
    }.map { |node|
      TemplateCell.new(node)
    }
  end

  def substitute_contents
    # ast.each do |template_call|
    #   key = cell.content.gsub('{{','').gsub('}}','')
    #   cell.content = FakeData.send(key.to_sym)
    # end
  end

  def ast
    @ast = []

    dynamic_cells.each do |cell|
      # Iterators go straight in (currently, this will change to allow
      # nesting)
      if cell.iterator?
        @ast << cell
      # Substitutions can be nested within iterators
      elsif cell.substitution?
        if @ast.last && @ast.last.iterator? && @ast.last.open?
          @ast.last.substitutions << cell
        else
          @ast << cell
        end
      # Iterators need to be closed
      elsif cell.closer?
        if @ast.last.iterator?
          @ast.last.closer = cell
        else
          raise 'No blocks open to close'
        end
      end
    end

    @ast
  end

  def write_tempfile(filepath)
    tempfile = File.open(filepath, 'w') do |f|
      f.write @xml.to_s
    end
  end

  # Duplicate the given 'from' cell to the 'to' cell - does not overwrite the
  # contents of the 'to' cell, but rather pushes that row and all others down
  # by one.
  #
  # options
  #   :from -
  #   :to   -
  #
  def duplicate_row(options)
    # for each row on and after the given from
    # increase the row number by 1
  end
end

class TemplateCell
  attr_accessor :substitutions, :closer, :node

  SUBSTITUTION = /^{{[_a-zA-Z]+}}$/
  ITERATOR     = /^{{#row for all (.*)+}}$/
  CLOSER       = /^{{#end}}$/

  def inspect
    preview = "Content: #{node.content}\n"

    if substitutions.any?
     preview = preview + "Subs: #{substitutions.inspect}"
    end

    unless closer.nil?
     preview = preview + "\nEnd:  #{closer.inspect}"
    end

    preview
  end

  def initialize(node)
    @node = node
    @substitutions = []
  end

  def tokens
    token_string = content.match(/{{(?<tokens>.*)}}/)[:tokens]
    token_string.split
  end

  def key
    tokens.first
  end

  def open?
    closer.nil?
  end

  def closed?
    ! open?
  end

  def content
    node.content.strip
  end

  def closer?
    content.match(CLOSER)
  end

  def iterator?
    content.match(ITERATOR)
  end

  def substitution?
    content.match(SUBSTITUTION)
  end
end
