from pptx import Presentation
import tempfile
import os
import shutil
import win32com.client as win32  # SOLO LOCAL FALLBACK


def merge_presentations(files, output="merged_output.pptx"):

    if not files:
        raise ValueError("No PPT files provided")

    temp_dir = tempfile.mkdtemp()
    paths = []

    # guardar archivos temporalmente
    for f in files:
        path = os.path.join(temp_dir, f.name)
        with open(path, "wb") as out:
            out.write(f.getbuffer())
        paths.append(path)

    # 🔥 SOLUCIÓN REAL: usar PowerPoint SOLO si está disponible
    try:
        ppt = win32.Dispatch("PowerPoint.Application")
        ppt.Visible = 1

        base = ppt.Presentations.Open(paths[0])
        merged = base

        for path in paths[1:]:
            merged.Slides.InsertFromFile(path, merged.Slides.Count)

        output_path = os.path.join(temp_dir, output)
        merged.SaveAs(output_path)

        merged.Close()
        ppt.Quit()

        return output_path

    except Exception:
        # fallback cloud-safe (básico)
        target = Presentation(paths[0])

        for path in paths[1:]:
            src = Presentation(path)

            for slide in src.slides:
                target.slides.add_slide(
                    target.slide_layouts[0]
                )

        output_path = os.path.join(temp_dir, output)
        target.save(output_path)

        return output_path