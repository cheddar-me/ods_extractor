class ODSExtractor::RowOutput
  def initialize(use_header_row:, &row_handler_block)
    @header_cells = []
    @row_handler = row_handler_block
    @use_header_row = use_header_row
  end

  def start_sheet(sheet_name)
    @sheet_name = sheet_name
    @header_cells = nil
  end

  def write_row(row_values)
    if @use_header_row
      if @header_cells
        @row_handler.call(sheet_name: @sheet_name, row: build_row_hash(row_values))
      else
        @header_cells = row_values.map(&:to_s)
      end
    else
      @row_handler.call(sheet_name: @sheet_name, row: row_values)
    end
  end

  def end_sheet
  end

  def build_row_hash(row_values)
    padded_row = row_values.take(@header_cells.length)
    (@header_cells.length - padded_row.length).times { padded_row << nil }
    @header_cells.zip(padded_row).to_h
  end
end
