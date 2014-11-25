
module Poi
  module Powerpoint

    def create_ppt(file = nil)
      file ? Rjb::import('org.apache.poi.xslf.usermodel.XMLSlideShow').new(poi_input_file(file)) : Rjb::import('org.apache.poi.xslf.usermodel.XMLSlideShow').new
    end

    def set_page_size(ppt, width, height)
      ppt.setPageSize(Rjb::import('java.awt.Dimension').new(height, width))
    end

    def save_ppt(ppt)
      out = poi_output_file('my_test.pptx')
      ppt.write(out)
      out.close()
    end

    def add_slide(ppt)
      ppt.createSlide()
    end

    def reorder_slides(ppt, options)
      return unless options
      options.each_key { |slide| ppt.setSlideOrder(slide, options[slide]) }
    end

    def add_photo_to_slide(workbook, sheet, row, column, image)
      picture_index = workbook.addPicture image, workbook.PICTURE_TYPE_JPEG
      drawing = workbook.getSheet(sheet).createDrawingPatriarch
      anchor = workbook.getCreationHelper.createClientAnchor
      anchor.setCol1 column
      anchor.setRow1 row
      drawing.createPicture(anchor, picture_index).resize
    end

    def create_spreadsheet(options, passed_styles = nil)
      workbook = create_excel_workbook
      if passed_styles
        styles = {}
        passed_styles.each do |style|
          styles[style.first] = create_cell_style workbook, style.last
        end
      end
      options.each do |sheet_hash|
        sheet = workbook.createSheet sheet_hash[:sheet][:name]
        sheet.setPrintGridlines(!!sheet_hash[:sheet][:print_grid_lines])
        sheet.setDisplayGridlines(!!sheet_hash[:sheet][:display_grid_lines])
        if sheet_hash[:sheet][:merged_regions]
          sheet_hash[:sheet][:merged_regions].each do |merged_region|
            sheet.addMergedRegion create_excel_cell_range_address.new(merged_region[:start_row], merged_region[:end_row],
                                                                      merged_region[:start_column], merged_region[:end_column])
          end
        end
        if sheet_hash[:sheet][:photos]
          sheet_hash[:sheet][:photos].each do |photo|
            add_photo_to_sheet workbook, sheet_hash[:sheet][:name], photo[:row], photo[:column], photo[:photo]
          end
        end
        if sheet_hash[:sheet][:column_widths]
          sheet_hash[:sheet][:column_widths].each do |column_width|
            sheet.setColumnWidth column_width.first, column_width.last
          end
        end
        if sheet_hash[:sheet][:row]
          sheet_hash[:sheet][:row].each do |row_hash|
            row = sheet.createRow row_hash[:row_index]
            row_hash[:row_height] ? row.setHeight(row_hash[:row_height]) : nil
            if row_hash[:cell]
              row_hash[:cell].each do |cell_hash|
                cell = row.createCell cell_hash[:cell_index]
                cell_hash[:value] ? cell.setCellValue(cell_hash[:value].to_s) : nil
                cell_hash[:style] ? cell.setCellStyle(styles[cell_hash[:style]]) : nil
              end
            end
          end
        end
      end
      workbook
    end
  end
end