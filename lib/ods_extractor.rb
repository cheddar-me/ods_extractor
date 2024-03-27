# frozen_string_literal: true

require_relative "ods_extractor/version"
require_relative "ods_extractor/csv_output"
require_relative "ods_extractor/sax_handler"
require_relative "ods_extractor/row_output"
require_relative "ods_extractor/sheet_filter_handler"

require "zip_kit"
require "nokogiri"

module ODSExtractor
  class Error < StandardError; end
  ACCEPT_ALL_SHEETS_PROC = ->(_sheet_name) { true }
  PROGRESS_HANDLER_PROC = ->(bytes_read, bytes_remaining) { true }
  CHUNK_SIZE = 32 * 1024

  def self.extract(input_io:, output_handler:, sheet_names: ACCEPT_ALL_SHEETS_PROC, progress_handler_proc: PROGRESS_HANDLER_PROC)
    # Feed the XML from the extractor directly to the SAX parser
    entries = ZipKit::FileReader.read_zip_structure(io: input_io)
    contentx_xml_zip_entry = entries.find { |e| e.filename == "content.xml" }

    raise Error, "No `content.xml` found in the ODS file" unless contentx_xml_zip_entry

    sax_handler = ODSExtractor::SAXHandler.new(output_handler)
    sax_filter = ODSExtractor::SheetFilterHandler.new(sax_handler, sheet_names)

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
    progress_handler_proc.call(0, contentx_xml_zip_entry.uncompressed_size)
    bytes_read = 0
    until ex.eof?
      chunk = ex.extract(CHUNK_SIZE)
      bytes_read += chunk.bytesize
      progress_handler_proc.call(bytes_read, contentx_xml_zip_entry.uncompressed_size - bytes_read)
      push_parser << chunk
    end
  ensure
    push_parser&.finish
  end
end
