# ODS Extractor

The ODS document contains multiple sheets, which we need to export. Of those sheets we fundamentally need just one,
but it is not the first one - and the soffice --headless command will only export the "selected" sheet to CSV. It is not
possible to switch to another sheet from the commandline - you need to create a macro and load it into openoffice,
and then trigger it from the commandline before exporting. Or you need to drive OO from UNO which you also need to install.

Here are some articles about solving this problem:

* https://www.linuxjournal.com/content/convert-spreadsheets-csv-files-python-and-pyuno-part-1v2
* https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=79869
* https://www.briankoponen.com/libreoffice-export-sheets-csv/
* https://ask.libreoffice.org/t/how-to-convert-specific-sheet-to-csv-via-command-line/11842
* https://askubuntu.com/questions/1042624/how-to-split-an-ods-spreadsheet-file-into-csv-files-per-sheet-on-the-terminal

etc etc etc

Now, that's all fine and dandy - but we can also solve this from the other end and just parse the ODS document ourselves.
This is easier than it might seem at first glance, but there are is a pitfall. There are gems which manipulate spreadsheet
documents in this way - but they first load the entire sheet set into memory. For example, as suggested here:
https://dipesh-prajapat.blogspot.com/ -  This won't fly for us because the sheets in
this doc are HUGE, and these gems are in fact not capable of dealing with a document that large. So we have to go a bit
more manual. An ODS document is just a ZIP with a huge XML inside of it (and whey I say huge I mean it - the ODS contains
a document of a whoopin' 430 MB of XML), both we can pry open, parse and extract the strings. For parsing the huge XML
we will use a SAX parser, which allows us to only unmarshal the tiny bits of XML we actually care about. We will also use
a zip extractor library so that we can unmarshal and extract at the same time, without a file in-between.

Because few people use this often: an XML SAX parser parses the document in a streaming fashion instead of
reconstructing a DOM tree. You can build a DOM tree based on a SAX parser but not the other way around. SAX parsers become
useful when parsing documents which are very large - and our "contents.xml" _is_ damn large indeed. To use a SAX parser we need
to write a handler, in the handler we are going to capture the elements we care about. The OpenOasis schema structures
a single sheet inside a "table:table" element, then every row is in "table:table-row", then every cell is within
"table:table-cell". Anything further down we can treat as text and just capture "as-is" (there are some wrapper tags
for paragraphs but these are not really important for our mission).
It is simpler than it seems, really.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ods_extractor'
```

And then execute:

    $ bundle install

Or install it yourself as:

    $ gem install ods_extractor

## Usage

TODO: Write usage instructions here

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/julik/ods_extractor. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/julik/ods_extractor/blob/master/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the ODSExtractor project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/julik/ods_extractor/blob/master/CODE_OF_CONDUCT.md).
