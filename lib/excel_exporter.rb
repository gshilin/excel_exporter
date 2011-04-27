class Array

  # Options:
  #
  # :only => only those columns will be exported
  # :except => all columns except these ones will be exported
  # Note: only one of the above options can be used at the same time
  #
  # :methods => list of columns to be exported that don't correspond to DB columns
  # :headers => Humanized name of column will be used as header by default;
  #             You can specify here a list of headers that must be used instead of default ones
  # :title   => name of workbook
  # :author  => name of author
  # :company => name of company
  #
  # Example:
  #
  #    list = Purchase.purchases_by_product(id).all
  #    fields = [:id, :username]
  #    headers = {:id => 'שובר', :username => 'שם'}
  #
  #    list.to_excel(
  #        :title => "#{Time.now.strftime("%d-%m-%Y")}_#{id}",
  #        :only => fields,
  #        :headers => headers,
  #        :types => {:product_id => 'Number', :user_id => 'Number', :price => 'Number'}
  #    )

  def to_excel(options = {})
    return '' if self.empty?

    klass = self.first.class
    attributes = self.first.attributes.keys.sort.map(&:to_sym)

    if options[:only]
      columns = Array(options[:only]) & attributes
    else
      columns = attributes - Array(options[:except])
    end

    columns += Array(options[:methods])

    return '' if columns.empty?

    time_now = Time.now

    header_names = options[:headers] || {}

    headers = columns.map { |column|
      name = (header_names && header_names[column]) ? header_names[column] : klass.human_attribute_name(column)
      "<Cell ss:StyleID='header'><Data ss:Type='String'>#{name}</Data></Cell>"
    }

    col_defs = columns.map { |column|
      '<Column ss:Width="141"/>'
    }

    types = options[:types] || {}
    
    body = self.map { |item|
      cols = columns.map { |colname|
        value = item.send(colname)
        type = colname == :id ? 'Number' : (types[colname] || 'String')
        "<Cell><Data ss:Type='#{type}'>#{value}</Data></Cell>"
      }
      "<Row>" + cols.join("\n") +"</Row>"
    }

    title = options[:title] || "#{time_now.strftime("%d-%m-%Y")}_#{options[:id]}"
    
    output = <<-OUTPUT
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>#{options[:author] || 'Unknown'}</Author>
  <LastAuthor>#{options[:author] || 'Unknown'}</LastAuthor>
  <Company>#{options[:company] || 'Unknown'}</Company>
  <Created>#{time_now.strftime('%Y-%m-%dT%H:%M:%SZ')}</Created>
  <LastSaved>#{time_now.strftime('%Y-%m-%dT%H:%M:%SZ')}</LastSaved>
 </DocumentProperties>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="header">
    <Alignment ss:Horizontal="Center"/>
    <Font ss:Size="8" ss:Bold="1"/>
    <Interior ss:Color="#d0d0d0" ss:Pattern="Solid"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="#{title}">
  <Table ss:ExpandedColumnCount="#{columns.length}" ss:ExpandedRowCount="#{self.length + 1}" x:FullColumns="1" x:FullRows="1" ss:DefaultRowHeight="15">
   #{ col_defs.join("\n") }
   <Row>#{ headers.join("\n") }</Row>
    #{ body.join("\n") }
  </Table>
 </Worksheet>
</Workbook>
    OUTPUT

    output
  end
end
	
