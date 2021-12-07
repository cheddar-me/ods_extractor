class ODSExtractor::CSVOutput
  def initialize(csv_directory_path)
    require "csv"
    @csv_directory_path = csv_directory_path
  end

  def start_sheet(sheet_name)
    @write_output_file = File.open(File.join(@csv_directory_path, sheet_name_to_csv_file_name(sheet_name)), "wb")
  end

  def write_row(row_values)
    @write_output_file << CSV.generate_line(row_values, force_quotes: true) # => "foo,0\n"
  end

  def end_sheet
    @write_output_file.close
  end

  def sheet_name_to_csv_file_name(sheet_name)
    # This is a subtle spot where there can be a security problem - we take an unsanitized sheet name
    # and we include it in a filesystem path. So some precaution needs to be taken.
    "#{sheet_name}.csv"
  end
end
