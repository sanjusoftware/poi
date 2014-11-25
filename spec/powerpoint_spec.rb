require 'spec_helper'
require 'poi/powerpoint'

describe 'POI::Powerpoint' do

  include Poi
  include Poi::Powerpoint

  context 'powerpoint' do

    before do
      initialize_poi
    end

    context 'create' do

      it 'should create a presentation with blank slide' do
        ppt = create_ppt
        expect(ppt).to_not be_nil
      end

      it 'should create a presentation with given file' do
        file = File.join(File.dirname(File.dirname(__FILE__)), 'spec', 'data', 'example_template.pptx')
        ppt = create_ppt(file)
        expect(ppt).to_not be_nil
      end

    end

    context 'editing' do

      it 'should set page size of a presentation' do
        ppt = create_ppt
        page_size = ppt.getPageSize()
        expect(page_size.width).to eq(720)
        expect(page_size.height).to eq(540)

        set_page_size(ppt, 1024, 768)

        page_size = ppt.getPageSize()
        expect(page_size.width).to eq(768)
        expect(page_size.height).to eq(1024)
      end

      it 'should save powerpoint' do
        ppt = create_ppt
        save_ppt(ppt)
        expect(Rjb::import('java.io.File').new('my_presentation.pptx')).not_to be_nil
        File.delete 'my_presentation.pptx'
      end

      it 'should add slide to a given powerpoint' do
        ppt = create_ppt
        expect(ppt.getSlides().count).to eq(0)
        add_slide(ppt)
        expect(ppt.getSlides().count).to eq(1)
      end

      it 'should reorder slides in the ppt' do
        file = File.join(File.dirname(File.dirname(__FILE__)), 'spec', 'data', 'example_template.pptx')
        ppt = create_ppt(file)
        slide1 = ppt.getSlides()[1]
        slide1_title = slide1.getTitle()
        slide2_title = ppt.getSlides()[2].getTitle()

        reorder_slides(ppt, {slide1 => 2})

        expect(ppt.getSlides()[1].getTitle()).to eq(slide2_title)
        expect(ppt.getSlides()[2].getTitle()).to eq(slide1_title)
      end

      it 'should add picture to the slide' do
        test_image = File.join(File.dirname(File.dirname(__FILE__)), 'spec', 'data', 'image001.jpg')
        ppt = create_ppt
        add_slide(ppt)
        add_photo_to_slide(ppt, 0, File.new(test_image).bytes.to_a)
        save_ppt(ppt, 'ppt_with_pic.pptx')
        ppt1 = create_ppt('ppt_with_pic.pptx')
        expect(File.new(test_image).bytes.to_a).to eq(ppt1.getAllPictures.get(0).getData.bytes.to_a)
        File.delete 'ppt_with_pic.pptx'
      end
    end

    context 'deleting' do
      it 'should delete the given slide from the ppt' do
        ppt = create_ppt
        add_slide(ppt)
        add_slide(ppt)
        expect(ppt.getSlides().count).to eq(2)
        delete_slide(ppt, 1)
        expect(ppt.getSlides().count).to eq(1)
      end

    end
  end

end