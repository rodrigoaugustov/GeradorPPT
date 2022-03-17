import copy


def change_pic(slide, old_picture, pic_file):
    if pic_file == "":
        old_picture._element.getparent().remove(old_picture._element)
        return None
    # Save the position and shape of old picture
    x, y, cx, cy = old_picture.left, old_picture.top, old_picture.width, old_picture.height

    # Add the new picture using old's picture shape and size
    slide.add_picture(pic_file, x, y, cx, cy)

    # Remove the old picture
    old_picture._element.getparent().remove(old_picture._element)
    return None


def ajusta_data(date):
    try:
        return date.date().strftime("%d/%m/%Y")
    except ValueError:
        return ''
    except AttributeError:
        return ''


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

    return dest


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
