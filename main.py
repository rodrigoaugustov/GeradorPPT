from pptx import Presentation
import copy


def change_pic(slide, old_picture, pic_file) -> None:
    """" Changes a picture in a specific slide
    :param Slide slide: The Slide object that contains the picture
    :param Shape old_picture: The Shape object of old picture
    :param str pic_file: The path to the new picture file
    """
    # Save the position and shape of old picture
    x, y, cx, cy = old_picture.left, old_picture.top, old_picture.width, old_picture.height

    # Add the new picture using old's picture shape and size
    slide.add_picture(pic_file, x, y, cx, cy)

    # Remove the old picture
    old_picture._element.getparent().remove(old_picture._element)
    return None


def clear_text(shape):
    """" Clear all texts from a shape """
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for run in p.runs:
                run.text = ""
    return shape


def change_text(shape, new_text):
    """" Insert text in a shape """
    if shape.has_text_frame:
        shape.text_frame.paragraphs[0].runs[0].text = new_text
    return shape


def duplicate_slide(pres, index):
    """ToDO"""
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

    return dest


def start():
    # Carrega a apresentação
    prs = Presentation('input.pptx')

    # Objeto slides
    slides = prs.slides

    # Objeto shapes
    shapes = slides[0].shapes

    # Any shape that contains texts may contain one or more paragraphs.
    # Each paragraphs may contains one or more RUNS.
    # If the paragraph have different formatting, each part will be an different RUN.
    # For example if the paragraphs have a bold word like "This is <b> an </b> example",
    # this paragraph will have three runs [This is, <b> an </n>, example].

    # So here we opt to clear all texts in all paragraphs and runs, and put the new text in only one paragraphs and run

    for shape in shapes:
        # Clear all paragraphs from a shape
        clear_text(shape)
        # Put a new text
        change_text(shape, 'NEW TEXT TO REPLACE')

    # Function to change picture. Arguments: Shapes of slide, The Shape that will be replaced, Path to new picture file
    change_pic(shapes, shapes[-1], 'FIGURE_FILE.png')

    # export pptx
    prs.save('OUTPUT.pptx')


if __name__ == '__main__':
    start()
