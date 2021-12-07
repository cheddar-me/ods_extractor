# ODS Extractor

This gem will help you export multiple sheets from OpenOffice Calc documents. The documents often contain multiple sheets,
and even though there is a command on OpenOffice to export a single sheet as CSV it will only export the sheet which
is currently selected. You will also need a full install of OpenOffice to do this (which, in a server environment, will mean
also a full GUI environment).

There _are_ ways  to use the `--headless` option to extract the sheets that you want, but you need to create a macro
and load it into openoffice, and then trigger it from the commandline before exporting.
Or you need to drive OO from UNO which you also need to install (see below in "Other solutions").

But instead of using a full OpenOffice/LibreOffice install we can also solve this from the other end and just parse the ODS document ourselves.
This is easier than it might seem at first glance, but there are is a pitfall. There are gems which manipulate spreadsheet
documents in this way - but they first load the entire document into memory, usually as a DOM. However this quickly breaks
down when huge documents are involved.

This gem uses a SAX parser to ingest the spreadsheets in a streaming fashion. An ODS document is just a ZIP with a huge XML inside
of it. For opening the ZIP file [zip_tricks](https://github.com/WeTransfer/zip_tricks) is used. For parsing the XML
a SAX parser from Nokogiri is used - you are likely to have Nokogiri already installed, as Rails uses it to sanitize HTML. The extraction
of the rows happens as the XML is fed to Nokogiri directly from the ZIP file, without having to create any intermediate files on
the file system.

## Usage

You need access to an IO object with random access which contains your ODS file. Imagine you want to capture all rows from the `Expenses` sheet
in your ODS, and you want to process them inline, as they get parsed:

```ruby
# The output receives the sheet name and the row contents. `use_header_rows`
# will yield `Hash` objects with headers from the first row as keys.
output = ODSExtractor::RowOutput.new(use_header_row: true) do |sheet_name:, row:|
  LineItem.create!(expenditure: row.fetch("Expenditure"), amount: row.fetch('Amount').to_i)
end

File.open(__dir__ + '/big_spreadsheet.ods') do |f|
  ODSExtractor.extract(input_io: f, output: out, sheet_names: "Expenses")
end
```

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ods_extractor'
```

And then execute:

    $ bundle install

Or install it yourself as:

    $ gem install ods_extractor

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/julik/ods_extractor. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/julik/ods_extractor/blob/master/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the ODSExtractor project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/julik/ods_extractor/blob/master/CODE_OF_CONDUCT.md).

## Other solutions

* https://www.linuxjournal.com/content/convert-spreadsheets-csv-files-python-and-pyuno-part-1v2
* https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=79869
* https://www.briankoponen.com/libreoffice-export-sheets-csv/
* https://ask.libreoffice.org/t/how-to-convert-specific-sheet-to-csv-via-command-line/11842
* https://askubuntu.com/questions/1042624/how-to-split-an-ods-spreadsheet-file-into-csv-files-per-sheet-on-the-terminal
