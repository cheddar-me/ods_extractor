require "delegate"

class ODSExtractor::SheetFilterHandler < SimpleDelegator

  def initialize(sax_handler, &block_for_sheet_name_filter)
    __setobj__(sax_handler)
    @sheet_name_filter_blk = block_for_sheet_name_filter
    @bypass = false
  end

  def start_element(name, attributes = [])
    return if @bypass

    if name == "table:table"
      sheet_name = attributes.to_h.fetch("table:name")
      if @sheet_name_filter_blk.call(sheet_name)
        @bypass = false
        super
      else
        $stderr.puts "Bypassing sheet #{sheet_name.inspect}"
        @bypass = true
        return
      end
    end

    super
  end

  def characters(string)
    return if @bypass
    super unless @bypass
  end

  def end_element(name)
    super unless @bypass
    @bypass = false if name == "table:table"
  end
end
