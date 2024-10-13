import os
import subprocess
from tqdm import tqdm  # For the progress bar

def convert_pptx_to_pdf(input_folder, output_folder):
    # Make sure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Get list of all .pptx files in the input folder
    pptx_files = [file for file in os.listdir(input_folder) if file.endswith(".pptx")]
    
    # Initialize the progress bar
    with tqdm(total=len(pptx_files), desc="Converting PPTX to PDF", unit="file") as pbar:
        # Iterate through all .pptx files and convert each to PDF
        for file_name in pptx_files:
            pptx_path = os.path.join(input_folder, file_name)
            pdf_file_name = file_name.replace(".pptx", ".pdf")
            pdf_path = os.path.join(output_folder, pdf_file_name)

            # Convert pptx to pdf using unoconv or soffice (LibreOffice)
            subprocess.run([
                'soffice', '--headless', '--convert-to', 'pdf', '--outdir', output_folder, pptx_path
            ], check=True)
            
            # Update progress bar after each conversion
            pbar.update(1)

if __name__ == "__main__":
    input_folder = r"/Users/amirmuz/code/python/pptx2pdf/pptx_to_convert/"  # Update with your input folder
    output_folder = r"/Users/amirmuz/code/python/pptx2pdf/converted/"  # Update with your output folder
    convert_pptx_to_pdf(input_folder, output_folder)
