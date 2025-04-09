import os
from tkinter import Tk, Label, Button, filedialog, StringVar, messagebox
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

def accessibility_processor(directory_path, new_directory):
    for filename in os.listdir(directory_path):
        if filename.endswith(".pptx"):
            filepath = os.path.join(directory_path, filename)
            prs = Presentation(filepath)

            filepath_alt_text = os.path.join(new_directory, filename)
            alt_text_output_file = f"{os.path.splitext(filepath_alt_text)[0]}_alt_text.txt"

            with open(alt_text_output_file, 'w', encoding='utf-8') as f:
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if "Title" in shape.name:
                            cursor_sp = slide.shapes[0]._element
                            cursor_sp.addprevious(shape._element)

                        if hasattr(shape, "alt_text") and shape.alt_text:
                            f.write(f" \n Shape: {shape.name} \n - Alt Text: {shape.alt_text} \n")
                        elif shape.shape_type == 13:  # Picture
                            alt_text = f"Image {shape.name}"  # Placeholder
                            f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n")
                        elif shape.has_text_frame:
                            alt_text = "Text content: " + shape.text
                            f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n")
                        else:
                            f.write(f" \n Shape: {shape.name} \n - No alt text available. \n")

            new_filepath = os.path.join(new_directory, f"{os.path.splitext(filename)[0]}_updated.pptx")
            prs.save(new_filepath)

            new_filepath_original_file = os.path.join(new_directory, filename)
            os.rename(filepath, new_filepath_original_file)

    messagebox.showinfo("Done", "Processing complete!")

def select_input_directory():
    folder = filedialog.askdirectory()
    if folder:
        input_dir.set(folder)

def select_output_directory():
    folder = filedialog.askdirectory()
    if folder:
        output_dir.set(folder)

def run_processor():
    if not input_dir.get() or not output_dir.get():
        messagebox.showerror("Error", "Please select both input and output directories.")
        return
    accessibility_processor(input_dir.get(), output_dir.get())


root = Tk()
root.title("PowerPoint Accessibility Processor")
root.geometry("500x250")

input_dir = StringVar()
output_dir = StringVar()

Label(root, text="Select directory with .pptx files:").pack(pady=5)
Button(root, text="Choose Input Directory", command=select_input_directory).pack()
Label(root, textvariable=input_dir, wraplength=400).pack()

Label(root, text="Select directory to save processed files:").pack(pady=10)
Button(root, text="Choose Output Directory", command=select_output_directory).pack()
Label(root, textvariable=output_dir, wraplength=400).pack()

Button(root, text="Run Processor", command=run_processor, bg="green", fg="white", padx=10, pady=5).pack(pady=20)

root.mainloop()
