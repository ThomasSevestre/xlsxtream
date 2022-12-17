# frozen_string_literal: true
require "date"
require "xlsxtream/core_extension"
require "xlsxtream/xml"

module Xlsxtream
  class Row

    ENCODING = Encoding.find('UTF-8')

    DATE_STYLE = 1
    TIME_STYLE = 2

    def initialize(row, rownum, options = {})
      @row = row
      @rownum = rownum
      @sst = options[:sst]
    end

    def to_xml
      column = String.new('A')
      xml = String.new(%Q{<row r="#{@rownum}">})

      @row.each do |value|
        unless value.nil?
          xml << value.to_xslx_value("#{column}#{@rownum}", @sst)
        end
        column.next!
      end

      xml << '</row>'
    end
  end
end
