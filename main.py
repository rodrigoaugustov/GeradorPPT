from pptx import Presentation
import copy

def change_pic(slide, old_picture, pic_file):
    # Save the position and shape of old picture
    x, y, cx, cy = old_picture.left, old_picture.top, old_picture.width, old_picture.height

    # Add the new picture using old's picture shape and size
    slide.add_picture(pic_file, x, y, cx, cy)

    # Remove the old picture
    old_picture._element.getparent().remove(old_picture._element)
    return None

# Em teste - Instavel
def duplicate_slide(pres, index):
    source = pres.slides[index]
    try:
        blank_slide_layout = pres.slide_layouts[6]
    except:
        blank_slide_layout = pres.slide_layouts[len(pres.slide_layouts) - 1]
    dest = pres.slides.add_slide(blank_slide_layout)

    for shp in source.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        dest.shapes._spTree.insert_element_before(newel, 'p:extLst')

    # for key, value in source.part.rels.items():
    #     if not "notesSlide" in value.reltype:
    #         if value.is_external:
    #             dest.part.rels.add_relationship(value.reltype, value.target_ref, value.rId, value.is_external)
    #         else:
    #             dest.part.rels.add_relationship(value.reltype, value.target_part, value.rId)

    return dest

def start():
    # Carrega a apresentação
    prs = Presentation('PPT_NAME.pptx')

    # duplica slides
    duplicate_slide(prs, 0)

    # Objeto slides
    slides = prs.slides

    # Objeto shapes
    shapes = slides[2].shapes

    # Function to change picture. Arguments: Shapes of slide, The Shape that will be replaced, Path to new picture file
    change_pic(shapes, shapes[-1], 'FIGURE_FILE.png')

    # export pptx
    prs.save('OUTPUT.pptx')

if __name__ == '__main__':
    start()
