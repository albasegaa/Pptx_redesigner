from pptx import Presentation
import tempfile
import os


def merge_presentations(files, output="merged_output.pptx"):

    if not files:
        raise ValueError("No PPT files provided")

    temp_dir = tempfile.mkdtemp()
    paths = []

    # guardar uploads en disco temporal
    for f in files:
        path = os.path.join(temp_dir, f.name)
        with open(path, "wb") as out:
            out.write(f.getbuffer())
        paths.append(path)

    # base presentation
    target = Presentation(paths[0])

    # añadir slides del resto
    for path in paths[1:]:
        src = Presentation(path)

        for slide in src.slides:

            layout = target.slide_layouts[6]
            new_slide = target.slides.add_slide(layout)

            for shape in slide.shapes:
                try:
                    new_slide.shapes._spTree.insert_element_before(
                        shape.element,
                        'p:extLst'
                    )
                except:
                    pass

    output_path = os.path.join(temp_dir, output)
    target.save(output_path)

    return output_path