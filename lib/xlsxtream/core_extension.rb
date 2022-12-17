class Object
  def to_xslx_value(cid, sst)
    to_s.to_xslx_value(cid, sst)
  end
end

class Numeric
  def to_xslx_value(cid, _)
    %Q{<c r="#{cid}" t="n"><v>#{self}</v></c>}
  end
end

class TrueClass
  def to_xslx_value(cid, _)
    %Q{<c r="#{cid}" t="b"><v>1</v></c>}
  end
end

class FalseClass
  def to_xslx_value(cid, _)
    %Q{<c r="#{cid}" t="b"><v>0</v></c>}
  end
end

class Time
  def to_xslx_value(cid, _)
    # Local dates are stored as UTC by truncating the offset:
    # 1970-01-01 00:00:00 +0200 => 1970-01-01 00:00:00 UTC
    # This is done because SpreadsheetML is not timezone aware.
    oa_date = (to_f + utc_offset) / 86400 + 25569

    %Q{<c r="#{cid}" s="#{Xlsxtream::Row::TIME_STYLE}"><v>#{oa_date}</v></c>}
  end
end

class DateTime
  def to_xslx_value(cid, _)
    _, jd, df, sf, of = marshal_dump
    oa_date = jd - 2415019 + (df + of + sf / 1e9) / 86400

     %Q{<c r="#{cid}" s="#{Xlsxtream::Row::TIME_STYLE}"><v>#{oa_date}</v></c>}
  end
end

class Date
  if RUBY_ENGINE == 'ruby'
    def to_xslx_value(cid, _)
      oa_date = (jd - 2415019).to_f
      %Q{<c r="#{cid}" s="#{Xlsxtream::Row::DATE_STYLE}"><v>#{oa_date}</v></c>}
    end
  else
    def to_xslx_value(cid, _)
      oa_date = jd - 2415019 + (hour * 3600 + sec + sec_fraction.to_f) / 86400
      %Q{<c r="#{cid}" s="#{Xlsxtream::Row::DATE_STYLE}"><v>#{oa_date}</v></c>}
    end
  end
end

class String
  def xlsx_shared_string
    Xlsxtream::SharedString.new(self)
  end

  def to_xslx_value(cid, sst)
    if empty?
      ""
    else
      if encoding != Xlsxtream::Row::ENCODING
        value = encode(Xlsxtream::Row::ENCODING)
      else
        value = self
      end

      %Q{<c r="#{cid}" t="inlineStr"><is><t>#{Xlsxtream::XML.escape_value(value)}</t></is></c>}
    end
  end
end

module Xlsxtream
  class SharedString < ::String
    def to_xslx_value(cid, sst)
      if empty?
        ""
      else
        if encoding != Xlsxtream::Row::ENCODING
          value = encode(Xlsxtream::Row::ENCODING)
        else
          value = self
        end

        %Q{<c r="#{cid}" t="s"><v>#{sst[value]}</v></c>}
      end
    end
  end
end
