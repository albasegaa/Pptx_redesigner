from pptx import Presentation
import tempfile
import os


def merge_presentations(files, output="merged_output.pptx"):

    if not files:
        raise ValueError("No PPT files provided")

    # convertir uploads a paths temporales
    temp_dir = tempfile.mkdtemp()
    paths = []

    for f in files:
        path = os.path.join(temp_dir, f.name)
        with open(path, "wb") as out:
            out.write(f.getbuffer())
        paths.append(path)

    # crear base
    merged = Presentation(paths[0])

    # añadir slides del resto
    for path in paths[1:]:
        prs = Presentation(path)

        for slide in prs.slides:
            layout = merged.slide_layouts[6]
            new_slide = merged.slides.add_slide(layout)

            for shape in slide.shapes:
                new_slide.shapes._spTree.insert_element_before(
                    shape.element,
                    'p:extLst'
                )

    output_path = os.path.join(temp_dir, output)
    merged.save(output_path)

    return output_path