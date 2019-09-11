require 'zip/filesystem'
require 'nokogiri'

module Dullard; end

class Dullard::Workbook
  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  # Some additional formats added by Paul Hendryx (phendryx@gmail.com) that are common in LibreOffice.
  FORMATS = {
    'general' => :float,
    '0' => :float,
    '0.00' => :float,
    '#,##0' => :float,
    '#,##0.00' => :float,
    '0%' => :percentage,
    '0.00%' => :percentage,
    '0.00E+00' => :float,
    '# ?/?' => :float, #??? TODO:
    '# ??/??' => :float, #??? TODO:
    'mm-dd-yy' => :date,
    'd-mmm-yy' => :date,
    'd-mmm' => :date,
    'mmm-yy' => :date,
    'h:mm am/pm' => :time,
    'h:mm:ss am/pm' => :time,
    'h:mm' => :time,
    'h:mm:ss' => :time,
    'm/d/yy h:mm' => :datetime,
    '#,##0 ;(#,##0)' => :float,
    '#,##0 ;[red](#,##0)' => :float,
    '#,##0.00;(#,##0.00)' => :float,
    '#,##0.00;[red](#,##0.00)' => :float,
	'#,##0_);[Red]($#,##0)' => :float, 			# Dan
	'#,##0.00_);[Red]($#,##0.00)' => :float, 	# Dan
    'mm:ss' => :time,
    '[h]:mm:ss' => :time,
    'mmss.0' => :time,
    '##0.0e+0' => :float,
    '@' => :float,
    #-- zusaetzliche Formate, die nicht standardmaessig definiert sind:
    "yyyy\\-mm\\-dd" => :date,
    'dd/mm/yy' => :date,
    'hh:mm:ss' => :time,
    "dd/mm/yy\\ hh:mm" => :datetime,
    'm/d/yy' => :date,
    'mm/dd/yy' => :date,
    'mm/dd/yyyy' => :date,
  }

  STANDARD_FORMATS = {
    0 => 'General',
    1 => '0',
    2 => '0.00',
    3 => '#,##0',
    4 => '#,##0.00',
    9 => '0%',
    10 => '0.00%',
    11 => '0.00E+00',
    12 => '# ?/?',
    13 => '# ??/??',
    14 => 'mm-dd-yy',
    15 => 'd-mmm-yy',
    16 => 'd-mmm',
    17 => 'mmm-yy',
    18 => 'h:mm AM/PM',
    19 => 'h:mm:ss AM/PM',
    20 => 'h:mm',
    21 => 'h:mm:ss',
    22 => 'm/d/yy h:mm',
    37 => '#,##0 ;(#,##0)',
    38 => '#,##0 ;[Red](#,##0)',
    39 => '#,##0.00;(#,##0.00)',
    40 => '#,##0.00;[Red](#,##0.00)',
    45 => 'mm:ss',
    46 => '[h]:mm:ss',
    47 => 'mmss.0',
    48 => '##0.0E+0',
    49 => '@',
  }

  COLOR_SCHEMES = {
    0 => nil,
    1 => nil,
    2 => '1F497D',
    3 => 'EEECE1',
    4 => '4F81BD',
    5 => 'C0504D',
    6 => '9BBB59',
    7 => '8064A2',
    8 => '4BACC6',
    9 => 'F79646',
    10 => '0000FF',
    11 => '800080'
  }

  # Source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.indexedcolors%28v=office.14%29.aspx
  COLOR_INDEXED = {
    0 => '000000',
    1 => 'FFFFFF',
    2 => 'FF0000',
    3 => '00FF00',
    4 => '0000FF',
    5 => 'FFFF00',
    6 => 'FF00FF',
    7 => '00FFFF',
    8 => '000000',
    9 => 'FFFFFF',
    10 => 'FF0000',
    11 => '00FF00',
    12 => '0000FF',
    13 => 'FFFF00',
    14 => 'FF00FF',
    15 => '00FFFF',
    16 => '800000',
    17 => '008000',
    18 => '000080',
    19 => '808000',
    20 => '800080',
    21 => '008080',
    22 => 'C0C0C0',
    23 => '808080',
    24 => '9999FF',
    25 => '993366',
    26 => 'FFFFCC',
    27 => 'CCFFFF',
    28 => '660066',
    29 => 'FF8080',
    30 => '0066CC',
    31 => 'CCCCFF',
    32 => '000080',
    33 => 'FF00FF',
    34 => 'FFFF00',
    35 => '00FFFF',
    36 => '800080',
    37 => '800000',
    38 => '008080',
    39 => '0000FF',
    40 => '00CCFF',
    41 => 'CCFFFF',
    42 => 'CCFFCC',
    43 => 'FFFF99',
    44 => '99CCFF',
    45 => 'FF99CC',
    46 => 'CC99FF',
    47 => 'FFCC99',
    48 => '3366FF',
    49 => '33CCCC',
    50 => '99CC00',
    51 => 'FFCC00',
    52 => 'FF9900',
    53 => 'FF6600',
    54 => '666699',
    55 => '969696',
    56 => '003366',
    57 => '339966',
    58 => '003300',
    59 => '333300',
    60 => '993300',
    61 => '993366',
    62 => '333399',
    63 => '333333',
    64 => nil,
    65 => nil
  }

  def initialize(file, user_defined_formats = {}, include_formatting = true)
    @file = file
    @zipfs = Zip::File.open(@file)
    @user_defined_formats = user_defined_formats
    @formatting = include_formatting
    read_styles(include_formatting)
  end

  def sheets
    workbook = Nokogiri::XML::Document.parse(@zipfs.file.open("xl/workbook.xml"))
    @sheets = workbook.css("sheet").each_with_index.map {|n,i| Dullard::Sheet.new(self, n.attr("name"), n.attr("sheetId"), i+1) }
  end

  def string_table
    @string_table ||= read_string_table
  end

  def read_string_table
    @string_table = []
    entry = ''
    Nokogiri::XML::Reader(@zipfs.file.open("xl/sharedStrings.xml")).each do |node|
      if node.name == "si" and node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
        entry = ''
      elsif node.name == "si" and node.node_type == Nokogiri::XML::Reader::TYPE_END_ELEMENT
        @string_table << entry
      elsif node.value?
        entry << node.value
      end
    end
    @string_table
  end

  def has_formatting?
    @formatting
  end

  def read_styles(include_formatting)
    doc = Nokogiri::XML(@zipfs.file.open("xl/styles.xml"))

    @num_formats = {}
    if include_formatting
      @font_formats = {}
      @fill_formats = {}
      @border_formats = {}
    end
    @cell_xfs = []

    doc.css('/styleSheet/numFmts/numFmt').each do |numFmt|
      numFmtId = numFmt.attributes['numFmtId'].value.to_i
      formatCode = numFmt.attributes['formatCode'].value
      @num_formats[numFmtId] = formatCode
    end

    if include_formatting
      doc.css('/styleSheet/fonts/font').each_with_index do |font, i|
        @font_formats[i] = {
          b: !font.css('/b').empty?,
          i: !font.css('/i').empty?,
          u: !font.css('/u').empty?,
          sz: font.css('/sz').empty? ? nil : font.css('/sz').first.attributes['val'].value.to_i,
          name: font.css('/name').empty? ? nil : font.css('/name').first.attributes['val'].value,
          color: font.css('/color').empty? ? nil : node2color(font.css('/color').first)
        }
      end
      # TODO: only accepts "none" and "solid" pattern types, and as a result, no support for "bgColor" attibute
      doc.css('/styleSheet/fills/fill/patternFill').each_with_index do |fill, i|
        @fill_formats[i] = {
          type: fill.attributes['patternType'].value.to_s == 'none' || fill.attributes['patternType'].value.to_s == 'solid' ? fill.attributes['patternType'].value : 'solid',
          fgColor: fill.css('/fgColor').empty? ? nil : node2color(fill.css('/fgColor').first)
          #bgColor: fill.css('/bgColor').empty? ? nil : node2color(fill.css('/bgColor').first)
        }
      end
      # TODO: doesn't handle diagonal borders, doesn't accept colors yet
      doc.css('/styleSheet/borders/border').each_with_index do |border, i|
        leftBorder = border.css('/left').first
        rightBorder = border.css('/right').first
        topBorder = border.css('/top').first
        bottomBorder = border.css('/bottom').first
        @border_formats[i] = {
          left: leftBorder.attributes.length == 0 ? nil : leftBorder.attributes['style'].value,
          right: rightBorder.attributes.length == 0 ? nil : rightBorder.attributes['style'].value,
          top: topBorder.attributes.length == 0 ? nil : topBorder.attributes['style'].value,
          bottom: bottomBorder.attributes.length == 0 ? nil : bottomBorder.attributes['style'].value
        }
      end
    end
    doc.css('/styleSheet/cellXfs/xf').each do |xf|
      # Alignment and word wrap
      if include_formatting
        alignment = {}
        alignments = xf.css('/alignment')
        if !alignments.empty?
          if alignments.first.attributes.has_key?('horizontal') then alignment['horizontal'] = alignments.first.attributes['horizontal'].value; end
          if alignments.first.attributes.has_key?('vertical')   then alignment['vertical'] = alignments.first.attributes['vertical'].value; end
          if alignments.first.attributes.has_key?('wrapText')   then alignment['wrapText'] = true; end
        end
      end

      # Generate format object
      @cell_xfs << (!include_formatting \
      ? { numFmtId: xf.attributes['numFmtId'].value.to_i }
      : {
        numFmtId: xf.attributes['numFmtId'].value.to_i,
        fontId: xf.attributes['fontId'].value.to_i,
        fillId: xf.attributes['fillId'].value.to_i,
        borderId: xf.attributes['borderId'].value.to_i,
        alignment: alignments.empty? ? nil : alignment
      })
    end
  end

  def external_links
    @external_links || read_external_links
  end

  def read_external_links
    @external_links = []

    # Loop over all files located in xl/externalLinks/_rels and extract the filenames
    # Code borrowed from http://www.rubydoc.info/github/rubyzip/rubyzip/master/Zip/FileSystem/ZipFsDir#foreach-instance_method
    path = "xl/externalLinks/_rels/"
    subDirEntriesRegex = Regexp.new("^#{path}([^/]+)$")
    @zipfs.each do |fileName|
      match = subDirEntriesRegex.match(fileName.to_s)
      if !match.nil?
        doc = Nokogiri::XML(@zipfs.file.open(fileName.to_s))
        doc.css('/Relationships/Relationship').each do |external_link|
          @external_links << { :id => match[1].reverse[9].to_i, :target => external_link.attributes['Target'].value }
        end
      end
    end

    @external_links
  end

  def has_macros?
     return @has_macros if defined?(@has_macros)
     return read_macros
  end

  # Added by Dan. Discover macros simply by searching for *.bin files in the root xl directory.
  # Note we could have also found linked macros by looping through xl/_rels/workbook.xml.rels and pulling out all entries of type = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
  # Not sure if remote macros can be linked in - e.g. if workbook.xml.rels could have macro entries while the .XLSM file does NOT actually contain *.bin files - worth checking on this
  def read_macros
      @has_macros = false
      binRegex = Regexp.new("^xl/([^/]+)\\.bin$")
      @zipfs.each do |filename|
          match = binRegex.match(fileName.to_s)
          if !match.nil?
              @has_macros = true
              return @has_macros
          end
      end

      @has_macros
  end

  # Added by Dan. Discover "drawings" - inserted objects like pictures, shapes
  def drawings
	@drawings || read_drawings
  end

  def read_drawings
	  @drawing_links = []
	  @drawings = []

	  # First, loop over all files located in xl/drawings/_rels and extract filenames of externally stored media (e.g., pictures/images)
      path = "xl/drawings/_rels/"
	  subDirEntriesRegex = Regexp.new("^#{path}([^/]+)$")
	  @zipfs.each do |fileName|
		match = subDirEntriesRegex.match(fileName.to_s)
		if !match.nil?
		  doc = Nokogiri::XML(@zipfs.file.open(fileName.to_s))
		  doc.css('/Relationships/Relationship').each do |drawing|
			target = drawing.attributes['Target'].value
			@drawing_links << { :sheet_id => match[1].reverse[9].to_i, :rel_id => drawing.attributes['Id'].value, :target => target, :file => @zipfs.file.open("xl/drawings/#{target}") }
		  end
		end
	  end

      # Second, loop over all XML files located in xl/drawings and extract embedded images details
	  # NOTE: "xdr:"s in xml documents are interpreted as namespaces, so use namespace separator - the pipe char - during css() searches
      path = "xl/drawings/"
      subDirEntriesRegex = Regexp.new("^#{path}([^/]+)\\.xml$")
      @zipfs.each do |fileName|
        match = subDirEntriesRegex.match(fileName.to_s)
        if !match.nil?
          doc = Nokogiri::XML(@zipfs.file.open(fileName.to_s))
          doc.css('/xdr|wsDr/xdr|twoCellAnchor').each do |two_cell_anchor|
			  from_node = two_cell_anchor.css('/xdr|from')
			  to_node = two_cell_anchor.css('/xdr|to')
			  pic_node = two_cell_anchor.css('/xdr|pic')
			  shape_node = two_cell_anchor.css('/xdr|sp')
			  type = !pic_node.empty? ? :pic : (!shape_node.empty? ? :shape : :other)

			  #  link pics to File objects of stored media, extracted in the @drawing_links loop above
			  media_link = nil
			  if type == :pic
				  rel_id = pic_node.css('/xdr|blipFill/a|blip').first.attributes['embed'].value
				  media_link = @drawing_links.find { |link| link[:sheet_id] == match[1].reverse[0].to_i && link[:rel_id] == rel_id }[:file]
			  end

			  # Finally, populate the @drawings object
			  # NOTE: rows and cols are zero-based when it comes to drawing/"xdr:..." attributes, so add 1
			  # NOTE: convert all xdr:rowOff and xdr:colOff (measured in "EMUs" or English Metric Units) to pixels, using calculation found here: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
			  @drawings << {
				  sheet_index: match[1].reverse[0].to_i,
				  from: {
					  col: from_node.css('/xdr|col').first.text.to_i + 1,
					  col_offset: from_node.css('/xdr|colOff').first.text.to_f / 914400 * 72,
					  row: from_node.css('/xdr|row').first.text.to_i + 1,
					  row_offset: from_node.css('/xdr|rowOff').first.text.to_f / 914400 * 72 },
			   	  to: {
					  col: to_node.css('/xdr|col').first.text.to_i + 1,
					  col_offset: to_node.css('/xdr|colOff').first.text.to_f / 914400 * 72,
					  row: to_node.css('/xdr|row').first.text.to_i + 1,
					  row_offset: to_node.css('/xdr|rowOff').first.text.to_f / 914400 * 72 },
				  type: type,
				  media: media_link
			  }
          end
        end
      end

      @drawings
  end

  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  # convert internal excelx attribute to a format
  def attribute2format(s)
    id = @cell_xfs[s.to_i][:numFmtId].to_i
    result = @num_formats[id]

    if result == nil
      if STANDARD_FORMATS.has_key? id
        result = STANDARD_FORMATS[id]
      end
    end

    result.downcase
  end

  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  # Updated by Dan Adler
  def format2type(format)
    if FORMATS.has_key? format
      FORMATS[format]
    elsif @user_defined_formats.has_key? format
      @user_defined_formats[format]
    else
      # Previously, just return :float here...
      # Instead, updating to correctly identify percents, dates/times from numeric formats
      # Step 1, remove all quoted (i.e., displayed as non-replaced static text) sections
	  # 		AND all color and language bracket marker sections
      adj_format = format.gsub(/\".*?\"/, "").gsub(/\[.*?\]/, "")
      # Step 2, check if a percent, s date, a datetime, or a time
	  if adj_format.include?("%")
		:percentage
      elsif (adj_format.include?("y") || adj_format.include?("d") || adj_format.include?("m")) && !(adj_format.include?("h") || adj_format.include?("s"))
        :date
      elsif (adj_format.include?("y") || adj_format.include?("d") || adj_format.include?("mmm")) && (adj_format.include?("h") || adj_format.include?("s"))
        :datetime
      elsif !(adj_format.include?("y") || adj_format.include?("d") || adj_format.include?("mmm")) && (adj_format.include?("h") || adj_format.include?("s"))
        :time
      else
        :float
      end
    end
  end

  def attribute2FontFmt(s)
    id = @cell_xfs[s.to_i][:fontId].to_i
    return id == 0 ? nil : @font_formats[id]
  end

  def attribute2FillFmt(s)
    id = @cell_xfs[s.to_i][:fillId].to_i
    return id == 0 ? nil : @fill_formats[id]
  end

  def attribute2BorderFmt(s)
    id = @cell_xfs[s.to_i][:borderId].to_i
    return id == 0 ? nil : @border_formats[id]
  end

  def attribute2Alignment(s)
	   # nil if xf didn't have an <alignment> child
	  return @cell_xfs[s.to_i][:alignment].nil? ? nil : @cell_xfs[s.to_i][:alignment].symbolize_keys
  end

  def node2color(color_node)
    if color_node.attributes.has_key? 'theme'
      COLOR_SCHEMES[color_node.attributes['theme'].value.to_i]
    elsif color_node.attributes.has_key? 'rgb'
      color_node.attributes['rgb'].value[2..-1]
    elsif color_node.attributes.has_key? 'indexed'
      COLOR_INDEXED[color_node.attributes['indexed'].value.to_i]
    else
      nil
    end
  end

  def zipfs
    @zipfs
  end

  def close
    @zipfs.close
  end
end

class Dullard::Sheet
  attr_reader :name, :workbook
  def initialize(workbook, name, id, index)
    @workbook = workbook
    @name = name
    @id = id
    @index = index
    @file = @workbook.zipfs.file.open(path) if @workbook.zipfs.file.exist?(path)
    @shared_formulas = []
  end

  def string_lookup(i)
    @workbook.string_table[i]
  end

  def rows
    Enumerator.new(row_count) do |y|
      next unless @file
      @file.rewind

      shared = false
      shared_formula = false
      formula_value = nil
      row = { :cells => nil }
      column = nil
      cell_type = nil
      node_value_type = nil
      row_num = 0

      Nokogiri::XML::Reader(@file).each do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when "row"
            row[:cells] = []
            column = 0
            row_num += 1

            # If sheet skips past rows, yield empty ones
            rrow = node.attributes["r"]
            if rrow
              while rrow.to_i > row_num
                y.yield({ :cells => [] })
                row_num += 1
              end
            end

            # If this is an empty row itself (no child nodes), yield an empty one
            if node.empty_element?
              y.yield({ :cells => [] })
            end

            # NOTE: do this in separate methods below instead of aggregating here. No easy way to aggregate col widths in this loop
            # # If row has a custom height set, record it in the row object
            # if node.attributes["customHeight"] && node.attributes["customHeight"] == "1" && node.attributes["ht"]
            #   row[:height] = node.attributes["ht"]
            # end

            next
          when "c"
            rcolumn = node.attributes["r"]
            if rcolumn
              rcolumn.delete!("0-9")
              while column < self.class.column_names.size and rcolumn != self.class.column_names[column]
                row[:cells] << nil
                column += 1
              end
            end

            row[:cells] << {c: column, v: nil, f: nil, type: nil}

            if node.attributes.has_key?('s') && (@workbook.has_formatting? || (node.attributes['t'] != 's' && node.attributes['t'] != 'b'))
              cell_format_index = node.attributes['s'].to_i
              cell_num_format = @workbook.attribute2format(cell_format_index)
              cell_type = @workbook.format2type(cell_num_format)

              # Value alone not sufficient to determine type - dates/times generally appear as numbers (in the seconds since whatever date format), so need to store this specifically
              row[:cells].last[:type] = cell_type

              if @workbook.has_formatting?
                row[:cells].last[:num_format] = cell_num_format
                row[:cells].last[:font] = @workbook.attribute2FontFmt(cell_format_index)
                row[:cells].last[:fill] = @workbook.attribute2FillFmt(cell_format_index)
                row[:cells].last[:border] = @workbook.attribute2BorderFmt(cell_format_index)
                row[:cells].last[:alignment] = @workbook.attribute2Alignment(cell_format_index)
              end
            else
              if @workbook.has_formatting?
                row[:cells].last[:num_format] = nil
                row[:cells].last[:font] = nil
                row[:cells].last[:fill] = nil
                row[:cells].last[:border] = nil
                row[:cells].last[:alignment] = nil
              end
            end

            shared = node.attribute("t") == "s"
            column += 1
            next
          when "v"
            node_value_type = "v"
            next
          when "f"
            shared_formula = node.attribute("t") == "shared"
            formula_value = nil

            if shared_formula
              si = node.attribute("si").to_i
              if @shared_formulas[si]
                formula_value = generate_formula(row_num, column, @shared_formulas[si])
              else
                formula_value = Nokogiri::XML.fragment(node.inner_xml).text
                @shared_formulas << {
                  row: row_num,
                  col: column,
                  f: formula_value
                }
              end

              # If shared, formula may not have a "child" node containing the formula itself. Instead, record it here
              row[:cells].last[:f] = formula_value
            end

            node_value_type = "f"
            next
          when "t"
            node_value_type = "t"
            next
          end
        when Nokogiri::XML::Reader::TYPE_END_ELEMENT
          if node.name == "row"
            y.yield row
            next
          elsif node.name == "sheetData"
            # NOTE: This is where we exit the XML loop - if we reach the end of the enclosing sheetData element, we know that there are no more cells/values/formulas/etc left to parse
            #   Instead, there are some follow-up tags - such as conditionalFormatting, cfRule, mergeCells/mergeCell, and formula elements
            #   While we exit the block iterator here, the last row was already yielded by the previous iteration (when it saw a node of type Nokogiri::XML::Reader::TYPE_END_ELEMENT with name "row")
            # TODO: To add conditional formatting or merged cell parsing, DON'T exit here, and instead handle cfRule and formula or mergeCell nodes in the Nokogiri::XML::Reader::TYPE_ELEMENT case above
            #   (as well as the text nodes that come within them - which could complicate the "t" and "v" handling below)
            break
          end
        end

        value = node.value

        if value
          if node_value_type == 'v' || node_value_type == 't'
            case cell_type
              when :datetime, :time, :date
                value = (DateTime.new(1899,12,30) + value.to_f)
              when :percentage # Do nothing
              when :float
                value = value.to_f
              else
                # leave as string
            end
            cell_type = nil

            row[:cells].last[:v] = (shared ? string_lookup(value.to_i) : value)
          elsif node_value_type == 'f'
            # If it's a shared formula, this will never be called, since there's no "node" containing the formula string
            row[:cells].last[:f] = value
          end

          node_value_type = nil
        end
      end
    end
  end

  # Returns A to ZZZ.
  def self.column_names
    if @column_names
      @column_names
    else
      proc = Proc.new do |prev|
        ("#{prev}A".."#{prev}Z").to_a
      end
      x = proc.call("")
      y = x.map(&proc).flatten
      z = y.map(&proc).flatten
      @column_names = x + y + z
    end
  end

  def row_count
    if defined? @row_count
      @row_count
    elsif @file
      @file.rewind
      Nokogiri::XML::Reader(@file).each do |node|
        if node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when "dimension"
            if ref = node.attributes["ref"]
              break @row_count = ref.scan(/\d+$/).first.to_i
            end
          when "sheetData"
            break @row_count = nil
          end
        end
      end
    end
  end

def drawings
	workbook.drawings.select { |drawing| drawing[:sheet_index] == @index }
end

def col_widths
    if @file
        @file.rewind
        widths = [nil]
        doc = Nokogiri::XML(@file)
        doc.css('/worksheet/cols/col').each do |col|
            start_col = col.attributes["min"].value.to_i
            end_col = col.attributes["max"].value.to_i
            if start_col && col.attributes["customWidth"] && col.attributes["customWidth"].value == "1"
                while start_col > widths.length
                    widths << nil
                end

                # TODO: Update this calculation using more precise character-to-inch-to-point-to-pixel calculations
                # Sadly, col widths in XLSX files not stored as pixels. They're stored as "# of default characters" wide
                # This means that we must make two conversions - characters-to-inch (using 100/9 here, default for a new file on my system, though apparently 8.43 was standard for many years?)
                # And then inches-to-pixel (using 72ppi here, which again is default for me - but that's really not correct, and depends on display)

                # NOTE: Calculation here matches the calculation in Cells/models/cell_sheet.rb for file downloading

                (end_col - start_col + 1).times { widths << (col.attributes["width"].value.to_f / (100.0 / 9.0) * 72.0).to_i }
            end
        end
        return widths
    end
end

def row_heights
    if @file
        @file.rewind
        heights = [nil]
        doc = Nokogiri::XML(@file)
        doc.css('/worksheet/sheetData/row').each do |row|
            row_num = row.attributes["r"].value.to_i
            if row_num && row.attributes["customHeight"] && row.attributes["customHeight"].value == "1"
                while row_num > heights.length
                    heights << nil
                end
                heights << row.attributes["ht"].value.to_i
            elsif row.attributes["hidden"]
                while row_num > heights.length
                    heights << nil
                end
                heights << 0
            end
        end
        return heights
    end
end

  def generate_formula(new_row, new_col, formula)
    row_diff = new_row - formula[:row]
    col_diff = new_col - formula[:col]

    return formula[:f]
      .gsub(/\$([[:upper:]]+)(\d+)/) {
        # Fixed column, variable row
        "$#{$1}#{$2.to_i + row_diff}"
      }.gsub(/(?<!\$)([[:upper:]]+)\$(\d+)/) {
        # Variable column, fixed row
        "#{self.class.column_names[self.class.column_names.index($1) + col_diff]}$#{$2}"
      }.gsub(/(?<!\$)([[:upper:]]+)(\d+)/) {
        # Variable column, variable row
        "#{self.class.column_names[self.class.column_names.index($1) + col_diff]}#{$2.to_i + row_diff}"
      }
  end

  private
  def path
    "xl/worksheets/sheet#{@index}.xml"
  end

end
