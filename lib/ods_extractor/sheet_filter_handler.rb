class ODSExtractor::SheetFilterHandler < Nokogiri::XML::SAX::Document
  def initialize(sax_handler, filter_definition)
    @handler = sax_handler
    @sheet_name_filter_proc = build_filter_proc(filter_definition)
    @bypass = false
  end

  def start_element(name, attributes = [])
    if name == "table:table"
      sheet_name = attributes.to_h.fetch("table:name")
      if @sheet_name_filter_proc.call(sheet_name)
        @bypass = false
        @handler.start_element(name, attributes)
      else
        @bypass = true
      end
    end

    @handler.start_element(name, attributes) unless @bypass
  end

  def error(string)
    raise ODSExtractor::Error, "XML parse error: #{string}"
  end

  def characters(string)
    return if @bypass
    @handler.characters(string)
  end

  def end_element(name)
    return if @bypass
    @handler.end_element(name)
  end

  def build_filter_proc(filter_definition)
    case filter_definition
    when String
      ->(sheet_name) { sheet_name == filter_definition }
    when Regexp
      ->(sheet_name) { sheet_name =~ filter_definition }
    when Array
      filter_definitions = filter_definition.map { |defn| build_filter_proc(defn) }
      ->(sheet_name) { filter_definitions.any? { |prc| prc.call(sheet_name) } }
    when Method, Proc
      filter_definition.to_proc # Return as is
    else
      "Sheet name filter must be an Array of String|Regexp|callable or a String|Regexp|callable"
    end
  end
end
