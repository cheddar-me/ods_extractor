# frozen_string_literal: true

RSpec.describe ODSExtractor do
  it "has a version number" do
    expect(ODSExtractor::VERSION).not_to be nil
  end

  it "parses a small ODS" do
    out = ODSExtractor::CSVOutput.new(__dir__)
    File.open(__dir__ + "/small_spreadsheet.ods") do |f|
      ODSExtractor.extract(input_io: f, output_handler: out)
    end
  end

  it "parses and yields rows with headers" do
    parsed_rows = []
    out = ODSExtractor::RowOutput.new(use_header_row: true) do |sheet_name:, row:|
      parsed_rows << {sheet_name: sheet_name, row: row}
    end
    File.open(__dir__ + "/small_spreadsheet.ods") do |f|
      ODSExtractor.extract(input_io: f, output_handler: out)
    end

    expect(parsed_rows).to eq([
      {sheet_name: "Second Sheet", row: {"Header B" => "Batman", "Header C" => "1"}},
      {sheet_name: "Second Sheet", row: {"Header B" => "Batman", "Header C" => "2"}},
      {sheet_name: "Первый лист", row: {"Header 1" => "na", "Header 2" => "na", "Header 3" => "na", "Header 4" => "na", "" => "overflow row"}},
      {sheet_name: "Первый лист", row: {"Header 1" => "na", "Header 2" => "na", "Header 3" => "na", "Header 4" => "na", "" => ""}},
      {sheet_name: "Первый лист", row: {"Header 1" => "na", "Header 2" => "na", "Header 3" => "", "Header 4" => "", "" => ""}},
      {sheet_name: "Первый лист", row: {"Header 1" => "na", "Header 2" => "na", "Header 3" => "", "Header 4" => "", "" => ""}},
      {sheet_name: "Первый лист", row: {"Header 1" => "Bat", "Header 2" => "man", "Header 3" => "", "Header 4" => "", "" => ""}},
      {sheet_name: "Первый лист", row: {"Header 1" => "Bat", "Header 2" => "maaan", "Header 3" => "", "Header 4" => "", "" => ""}},
      {sheet_name: "Первый лист", row: {"Header 1" => "Batman!", "Header 2" => "", "Header 3" => "", "Header 4" => "", "" => ""}},
      {sheet_name: "Первый лист", row: {"Header 1" => "space", "Header 2" => "space", "Header 3" => "", "Header 4" => "", "" => ""}},
      {sheet_name: "Первый лист", row: {"Header 1" => "space", "Header 2" => "", "Header 3" => "", "Header 4" => "", "" => ""}}
    ])
  end

  it "parses and yields rows as values" do
    parsed_rows = []
    out = ODSExtractor::RowOutput.new(use_header_row: false) do |sheet_name:, row:|
      parsed_rows << {sheet_name: sheet_name, row: row}
    end
    File.open(__dir__ + "/small_spreadsheet.ods") do |f|
      ODSExtractor.extract(input_io: f, output_handler: out)
    end

    expect(parsed_rows).to eq([
      {row: ["Header B", "Header C"], sheet_name: "Second Sheet"},
      {row: ["Batman", "1"], sheet_name: "Second Sheet"},
      {row: ["Batman", "2"], sheet_name: "Second Sheet"},
      {row: ["Header 1", "Header 2", "Header 3", "Header 4", ""], sheet_name: "Первый лист"},
      {row: ["na", "na", "na", "na", "overflow row"], sheet_name: "Первый лист"},
      {row: ["na", "na", "na", "na", ""], sheet_name: "Первый лист"},
      {row: ["na", "na", "", "", ""], sheet_name: "Первый лист"},
      {row: ["na", "na", "", "", ""], sheet_name: "Первый лист"},
      {row: ["Bat", "man", "", "", ""], sheet_name: "Первый лист"},
      {row: ["Bat", "maaan", "", "", ""], sheet_name: "Первый лист"},
      {row: ["Batman!", "", "", "", ""], sheet_name: "Первый лист"},
      {row: ["space", "space", "", "", ""], sheet_name: "Первый лист"},
      {row: ["space", "", "", "", ""], sheet_name: "Первый лист"}
    ])
  end

  it "applies the sheet name filter using a Proc" do
    parsed_rows = []
    out = ODSExtractor::RowOutput.new(use_header_row: false) do |sheet_name:, row:|
      parsed_rows << {sheet_name: sheet_name, row: row}
    end
    filter = ->(sheet_name) { sheet_name =~ /Первый/ }
    File.open(__dir__ + "/small_spreadsheet.ods") do |f|
      ODSExtractor.extract(input_io: f, output_handler: out, sheet_names: filter)
    end

    expect(parsed_rows).to eq([
      {row: ["Header 1", "Header 2", "Header 3", "Header 4", ""], sheet_name: "Первый лист"},
      {row: ["na", "na", "na", "na", "overflow row"], sheet_name: "Первый лист"},
      {row: ["na", "na", "na", "na", ""], sheet_name: "Первый лист"},
      {row: ["na", "na", "", "", ""], sheet_name: "Первый лист"},
      {row: ["na", "na", "", "", ""], sheet_name: "Первый лист"},
      {row: ["Bat", "man", "", "", ""], sheet_name: "Первый лист"},
      {row: ["Bat", "maaan", "", "", ""], sheet_name: "Первый лист"},
      {row: ["Batman!", "", "", "", ""], sheet_name: "Первый лист"},
      {row: ["space", "space", "", "", ""], sheet_name: "Первый лист"},
      {row: ["space", "", "", "", ""], sheet_name: "Первый лист"}
    ])
  end

  describe "applies the sheet name filters" do
    filters = [
      /Первый/,
      "Первый лист",
      ["Второй лист", "Первый лист"],
      ->(n) { n.start_with?("Пер") }
    ]
    filters.each do |filter|
      it "when the filter is defined as #{filter.class}" do
        parsed_rows = []
        out = ODSExtractor::RowOutput.new(use_header_row: false) do |sheet_name:, row:|
          parsed_rows << {sheet_name: sheet_name, row: row}
        end
        File.open(__dir__ + "/small_spreadsheet.ods") do |f|
          ODSExtractor.extract(input_io: f, output_handler: out, sheet_names: filter)
        end

        expect(parsed_rows).to eq([
          {row: ["Header 1", "Header 2", "Header 3", "Header 4", ""], sheet_name: "Первый лист"},
          {row: ["na", "na", "na", "na", "overflow row"], sheet_name: "Первый лист"},
          {row: ["na", "na", "na", "na", ""], sheet_name: "Первый лист"},
          {row: ["na", "na", "", "", ""], sheet_name: "Первый лист"},
          {row: ["na", "na", "", "", ""], sheet_name: "Первый лист"},
          {row: ["Bat", "man", "", "", ""], sheet_name: "Первый лист"},
          {row: ["Bat", "maaan", "", "", ""], sheet_name: "Первый лист"},
          {row: ["Batman!", "", "", "", ""], sheet_name: "Первый лист"},
          {row: ["space", "space", "", "", ""], sheet_name: "Первый лист"},
          {row: ["space", "", "", "", ""], sheet_name: "Первый лист"}
        ])
      end
    end
  end

  it "parses a large ODS" do
    out = ODSExtractor::CSVOutput.new(__dir__)
    progress_reports = []
    File.open(__dir__ + "/uk-sanctions-list.ods") do |f|
      progress = ->(read, remaining) { progress_reports << [read, remaining] }
      ODSExtractor.extract(input_io: f, output_handler: out, progress_handler_proc: progress)
    end

    expect(progress_reports[0]).to eq([0, 430309923])
    expect(progress_reports[-1]).to eq([430309923, 0])
  end
end
