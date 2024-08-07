import os
import subprocess
import pytesseract
from PIL import Image
import pandas as pd
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

# Set paths to executables
ffmpeg_path = r'C:\ffmpeg\bin\ffmpeg.exe'
tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# Country codes and expected mobile number lengths
country_data = {
    "+93": 9, "+355": 9, "+213": 9, "+376": 6, "+244": 9, "+54": 10,
    "+61": 9, "+43": 10, "+880": 10, "+32": 9, "+55": 10, "+1": 10,
    "+86": 11, "+20": 10, "+33": 9, "+49": 10, "+91": 10, "+62": 10,
    "+39": 10, "+81": 10, "+254": 9, "+52": 10, "+234": 10, "+92": 10,
    "+7": 10, "+966": 9, "+27": 9, "+82": 9, "+34": 9, "+90": 10,
    "+44": 10,
}

# Function to validate mobile numbers
def validate_mobile_number(number):
    # Remove spaces and hyphens
    number = re.sub(r"[ -]", "", number)
    # Check if number starts with a recognized country code
    for code, length in country_data.items():
        if number.startswith(code):
            # Validate length of mobile number without the country code
            if len(number[len(code):]) == length:
                return True
    return False

# Function to extract frames from video
def extract_frames(video_path, output_dir):
    command = f'"{ffmpeg_path}" -i "{video_path}" -vf fps=1 "{output_dir}/frame_%04d.png"'
    subprocess.run(command, shell=True, check=True)

# Function to extract phone numbers from text
def extract_phone_numbers(text):
    phone_regex = re.compile(r'\+?\d[\d\s\-\(\)]{7,}\d')
    return phone_regex.findall(text)

# Function to process an image
def process_image(image_path, phone_numbers, invalid_numbers):
    text = pytesseract.image_to_string(Image.open(image_path))
    extracted_numbers = extract_phone_numbers(text)
    for number in extracted_numbers:
        if validate_mobile_number(number):
            phone_numbers.add(number)
        else:
            invalid_numbers.add(number)

# Main function to process input and generate Excel file
def process_to_excel(input_paths, excel_output_path):
    try:
        frames_dir = './frames'
        os.makedirs(frames_dir, exist_ok=True)

        phone_numbers = set()
        invalid_numbers = set()

        messagebox.showinfo("Processing", "Starting the process...")

        for input_path in input_paths:
            if input_path.lower().endswith(('.mp4', '.avi', '.mov', '.mkv')):
                # If input is a video, extract frames
                messagebox.showinfo("Processing", f"Extracting frames from video: {os.path.basename(input_path)}...")
                extract_frames(input_path, frames_dir)
                for image_file in os.listdir(frames_dir):
                    if image_file.endswith('.png'):
                        image_path = os.path.join(frames_dir, image_file)
                        process_image(image_path, phone_numbers, invalid_numbers)
                # Clean up frames directory after processing each video
                shutil.rmtree(frames_dir)
                os.makedirs(frames_dir, exist_ok=True)
            elif input_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
                # If input is an image, process it directly
                messagebox.showinfo("Processing", f"Processing image: {os.path.basename(input_path)}...")
                process_image(image_path, phone_numbers, invalid_numbers)
            else:
                raise ValueError(f"Unsupported file format: {os.path.basename(input_path)}. Please provide a video or image file.")

        # Create DataFrame with valid and invalid numbers
        data = [(num, 'Valid') for num in phone_numbers] + [(num, 'Invalid') for num in invalid_numbers]
        df = pd.DataFrame(data, columns=['Contact Number', 'Validation Status'])
        df.to_excel(excel_output_path, index=False)

        if invalid_numbers:
            messagebox.showinfo("Process Completed", f"Contact numbers processed and saved to {excel_output_path}.\n\n"
                                                     f"{len(invalid_numbers)} numbers could not be validated:\n{', '.join(list(invalid_numbers)[:10])}...\n"
                                                     f"and {len(invalid_numbers) - 10} more.") if len(invalid_numbers) > 10 else messagebox.showinfo("Process Completed", 
                                                     f"Contact numbers processed and saved to {excel_output_path}.\n\n"
                                                     f"These numbers could not be validated: {', '.join(invalid_numbers)}.")
        else:
            messagebox.showinfo("Success", f"All contact numbers processed and saved to {excel_output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI functions
def select_input_files():
    file_paths = filedialog.askopenfilenames(filetypes=[
        ("Video and Image files", "*.mp4;*.avi;*.mov;*.mkv;*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, ';'.join(file_paths))

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, file_path)

def start_processing():
    input_paths = input_entry.get().split(';')
    output_path = output_entry.get()
    if input_paths and output_path:
        process_to_excel(input_paths, output_path)
    else:
        messagebox.showwarning("Input Error", "Please specify both input files and output Excel file.")

# Create the main window
root = tk.Tk()
root.title("EDUSOUL DATA CHECKER")

# Create UI elements
input_label = tk.Label(root, text="Select Input Files:")
input_label.pack(pady=5)

input_entry = tk.Entry(root, width=50)
input_entry.pack(pady=5)

input_button = tk.Button(root, text="Browse", command=select_input_files)
input_button.pack(pady=5)

output_label = tk.Label(root, text="Select Output Excel File:")
output_label.pack(pady=5)

output_entry = tk.Entry(root, width=50)
output_entry.pack(pady=5)

output_button = tk.Button(root, text="Browse", command=select_output_file)
output_button.pack(pady=5)

process_button = tk.Button(root, text="Start Processing", command=start_processing)
process_button.pack(pady=20)

# Run the application
root.mainloop()
