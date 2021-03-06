require "nokogiri"

# Because few people use this often: an XML SAX parser parses the document in a streaming fashion instead of
# reconstructing a DOM tree. You can build a DOM tree based on a SAX parser but not the other way around. SAX parsers become
# useful when parsing documents which are very large - and our "contents.xml" _is_ damn large indeed. To use a SAX parser we need
# to write a handler, in the handler we are going to capture the elements we care about. The OpenOasis schema structures
# a single sheet inside a "table:table" element, then every row is in "table:table-row", then every cell is within
# "table:table-cell". Anything further down we can treat as text and just capture "as-is" (there are some wrapper tags
# for paragraphs but these are not really important for our mission).
class ODSExtractor::SAXHandler < Nokogiri::XML::SAX::Document
  MAX_CELLS_PER_ROW = 2**14
  MAX_ROWS_PER_SHEET = 2**20

  def initialize(output_handler)
    @out = output_handler
  end

  def start_element(name, attributes = [])
    case name
      when "table:table"
        sheet_name = attributes.to_h.fetch("table:name")
        @out.start_sheet(sheet_name)
        @rows_output_so_far = 0
      when "table:table-row"
        # Here be dragons: https://stackoverflow.com/a/2741709/153886
        # Both rows and cells are actually _sparsely_ recorded in the XML, see below
        # for the same for cells.
        @row_repeats_n_times = attributes.to_h.fetch("table:number-rows-repeated", "1").to_i
        if @rows_output_so_far + @row_repeats_n_times >= MAX_ROWS_PER_SHEET
          # The ODS table contains "at most"
          # 1048576 rows. When we are at the last row, ODS will helpfully
          # tell us that there are "that many" repeat empty rows until the end of sheet.
          # These cells are useless for us of course, but if we repeat them literally
          # we will still output them to the CSV. We can use this to detect our last row.
          @row_repeats_n_times = 0
        end
        # and prepare an empty row
        @row = []
      when "table:table-cell"
        @cell_repeats_n_times = attributes.to_h.fetch("table:number-columns-repeated", "1").to_i
        if @row.length + @cell_repeats_n_times >= MAX_CELLS_PER_ROW
          # Again something pertinent: the ODS table contains "at most"
          # 2**14 columns - 16384. When we are at the last cell of the row, ODS will helpfully
          # tell us that there are "that many" repeat empty cells until the next row starts.
          # We can thus detect the last row by it having number-columns-repeated which creates N
          # similar cells. If we encounter that we can simply omit that cell, it is most certainly empty
          @cell_repeats_n_times = 0
        end
        @charbuf = String.new(capacity: 512) # Create a string which is unlikely to be resized all the time
    end
  end

  def characters(string)
    # @charbuf is only not-nil when we are inside a "table:table-cell" element, this allows us to skip
    # any chardata that is outside of cells for whatever reason
    @charbuf << string if @charbuf
  end

  def end_element(name)
    case name
      when "table:table"
        @out.end_sheet
      when "table:table-row"
        @rows_output_so_far += @row_repeats_n_times
        @row_repeats_n_times.times do
          @out.write_row(@row)
        end
      when "table:table-cell"
        @cell_repeats_n_times.times { @row << @charbuf.strip } # Have to strip due to XML having sometimes-significant whitespace
        @charbuf = nil
    end
  end
end
