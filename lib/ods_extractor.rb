# frozen_string_literal: true

require_relative "ods_extractor/version"
require_relative "ods_extractor/csv_output"
require_relative "ods_extractor/sax_handler"
require_relative "ods_extractor/sheet_filter_handler"
require 'zip_tricks'
require 'nokogiri'

module ODSExtractor
  class Error < StandardError; end
  TRUE_FN = ->(_sheet_name) { true }

  def self.extract(input_io:, output_handler:, sheet_name_filter_proc: TRUE_FN)
    # Feed the XML from the extractor directly to the SAX parser
    entries = ZipTricks::FileReader.read_zip_structure(io: input_io)
    contentx_xml_zip_entry = entries.find { |e| e.filename == "content.xml" }

    raise Error, "No `content.xml` found in the ODS file" unless contentx_xml_zip_entry

    sax_handler = ODSExtractor::SAXHandler.new(output_handler)
    sax_filter = ODSExtractor::SheetFilterHandler.new(sax_handler, &sheet_name_filter_proc)

    # Because we do not have a random access IO to the deflated XML inside the zip, but
    # we will be reading the deflated bytes and inflating them ourselves, we can't really
    # use the standard Parser - we need to use the PushParser. The Parser "reads" by itself
    # from the IO it has been given, PushParser can be fed bytes as we deflate them.
    push_parser = Nokogiri::XML::SAX::PushParser.new(sax_filter)

    # The "extract" call reads N bytes, inflates them and then returns them. We do not
    # know how big the inflated data will be before we inflate it, and the libxml2
    # push parser will abort with an error if we force-feed it chunks which are too big.
    # So read smol.
    ex = contentx_xml_zip_entry.extractor_from(input_io)
    bytes_read = 0
    yield(0, contentx_xml_zip_entry.uncompressed_size) if block_given?
    until ex.eof?
      chunk = ex.extract(64 * 1024)
      bytes_read += chunk.bytesize
      yield(bytes_read, contentx_xml_zip_entry.uncompressed_size - bytes_read) if block_given?
      push_parser << chunk
    end
    push_parser.finish
  end
end
