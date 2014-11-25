module Poi
  module Powerpoint

    def create_ppt(file = nil)
      file ? Rjb::import('org.apache.poi.xslf.usermodel.XMLSlideShow').new(poi_input_file(file)) : Rjb::import('org.apache.poi.xslf.usermodel.XMLSlideShow').new
    end

    def set_page_size(ppt, width, height)
      ppt.setPageSize(Rjb::import('java.awt.Dimension').new(height, width))
    end

    def save_ppt(ppt, file_path = nil)
      out = poi_output_file(file_path ? file_path : 'my_presentation.pptx')
      ppt.write(out)
      out.close()
    end

    def merge_ppts(ppts)
      merged_ppt = create_ppt
      ppts.each { |ppt| ppt.getSlides().each { |slide| merged_ppt.createSlide().importContent(slide) } }
      merged_ppt
    end

    def insert_ppt(ppt, ppt_to_insert, insert_at)
      merged_ppt = create_ppt
      ppt.getSlides().each_with_index do |slide, index|
        if index == insert_at
          merged_ppt = merge_ppts([merged_ppt, ppt_to_insert])
        end
        merged_ppt.createSlide().importContent(slide)
      end

      merged_ppt
    end

    def add_slide(ppt, title_text = nil, body_text = nil)
      if body_text || title_text
        slide_master = ppt.getSlideMasters()[0]
        slide_layout = slide_master.getLayout(Rjb::import('org.apache.poi.xslf.usermodel.SlideLayout').TITLE_AND_CONTENT)
        slide = ppt.createSlide(slide_layout)
        if body_text
          body = slide.getPlaceholder(1)
          body.clearText()
          textRun = body.addNewTextParagraph().addNewTextRun()
          textRun.setText(body_text)
        end

        if title_text
          title = slide.getPlaceholder(0)
          title.clearText()
          textRun = title.addNewTextParagraph().addNewTextRun()
          textRun.setText(title_text)
        end
        slide
      else
        ppt.createSlide()
      end
    end

    def delete_slide(ppt, slide_index)
      ppt.removeSlide(slide_index)
    end

    def reorder_slides(ppt, options)
      return unless options
      options.each_key { |slide| ppt.setSlideOrder(slide, options[slide]) }
    end

    def add_photo_to_slide(ppt, slide_index, image)
      picture_index = ppt.addPicture(image, Rjb::import('org.apache.poi.xslf.usermodel.XSLFPictureData').PICTURE_TYPE_PNG)
      slide = ppt.getSlides()[slide_index]
      slide.createPicture(picture_index)
    end

  end
end