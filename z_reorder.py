import os

from pptx import Presentation

from pptx.enum.shapes import MSO_SHAPE_TYPE

from alternat.generation import Generator


# alt text generation is not very good at the moment...

def accessibility_processor(directory):

    for filename in os.listdir(directory):

        if filename.endswith(".pptx"):

            filepath = os.path.join(directory, filename)

            prs = Presentation(filepath)


            filepath_alt_text = os.path.join(new_directory, filename)

            alt_text_output_file = f"{os.path.splitext(filepath_alt_text)[0]}_alt_text"

            

            # Iterate through all slides

            with open(alt_text_output_file, 'w', encoding='utf-8') as f:

                for slide in prs.slides:


                    # Iterate through all shapes in the slide

                    for shape in slide.shapes:

                        # Check if the shape has a name and if it's "Title" update zorder

                        if "Title" in shape.name:

                            cursor_sp = slide.shapes[0]._element

                            cursor_sp.addprevious(shape._element)


                        # Generate Alt Text

                        if hasattr(shape, "alt_text") and shape.alt_text:

                            f.write(f" \n Shape: {shape.name} \n - Alt Text: {shape.alt_text} \n")

                        elif shape.shape_type == 13:  # 13 is the shape type for pictures

                        # Generate alt text based on the image filename or content

                            alt_text = f"Image {shape.name}"  # Placeholder for actual alt text

                            f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n")

                        elif shape.has_text_frame:

                            # Generate alt text based on text content

                            alt_text = "Text content: " + shape.text

                            f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n")

                        else:

                            f.write(f" \n Shape: {shape.name} \n - No alt text available. \n")

            

            # Save the updated presentation with "_updated" suffix

            # Change this directory path

            new_filepath = os.path.join(new_directory, f"{os.path.splitext(filename)[0]}_updated.pptx")

            prs.save(new_filepath)


            #move original file to Processed folder

            new_filepath_original_file = os.path.join(new_directory, filename)

            os.rename(filepath, new_filepath_original_file)



# Change directory_path and new_directory to match local folders

directory_path = None

new_directory = None


accessibility_processor(directory_path)
