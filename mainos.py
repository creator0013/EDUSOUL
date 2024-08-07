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

# Function to extract frames from video
def extract_frames(video_path, output_dir):
    command = f'"{ffmpeg_path}" -i "{video_path}" -vf fps=1 "{output_dir}/frame_%04d.png"'
    subprocess.run(command, shell=True, check=True)

# Function to extract phone numbers from text
def extract_phone_numbers(text):
    phone_regex = re.compile(r'\+?\d[\d\s\-\(\)]{7,}\d')
    return phone_regex.findall(text)

# Function to process an image
def process_image(image_path, phone_numbers):
    text = pytesseract.image_to_string(Image.open(image_path))
    phone_numbers.update(extract_phone_numbers(text))

# Main function to process input and generate Excel file
def process_to_excel(input_path, excel_output_path):
    try:
        frames_dir = './frames'
        os.makedirs(frames_dir, exist_ok=True)

        phone_numbers = set()
        
        if input_path.lower().endswith(('.mp4', '.avi', '.mov', '.mkv')):
            # If input is a video, extract frames
            extract_frames(input_path, frames_dir)
            for image_file in os.listdir(frames_dir):
                if image_file.endswith('.png'):
                    image_path = os.path.join(frames_dir, image_file)
                    process_image(image_path, phone_numbers)
        elif input_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
            # If input is an image, process it directly
            process_image(input_path, phone_numbers)
        else:
            raise ValueError("Unsupported file format. Please provide a video or image file.")

        unique_numbers = sorted(phone_numbers)
        df = pd.DataFrame(unique_numbers, columns=['Contact Number'])
        df.to_excel(excel_output_path, index=False)

        print(f"Contact numbers extracted and saved to {excel_output_path}")
        messagebox.showinfo("Success", f"Process complete. Contacts saved to {excel_output_path}")

        # Clean up frames directory
        shutil.rmtree(frames_dir)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI functions
def select_input_file():
    file_path = filedialog.askopenfilename(filetypes=[
        ("Video and Image files", "*.mp4;*.avi;*.mov;*.mkv;*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, file_path)

def start_processing():
    input_path = input_entry.get()
    output_path = output_entry.get()
    if input_path and output_path:
        process_to_excel(input_path, output_path)
    else:
        messagebox.showwarning("Input Error", "Please specify both input file and output Excel file.")

# Create the main window
root = tk.Tk()
root.title("Media to Excel Converter")

# Create UI elements
input_label = tk.Label(root, text="Select Input File:")
input_label.pack(pady=5)

input_entry = tk.Entry(root, width=50)
input_entry.pack(pady=5)

input_button = tk.Button(root, text="Browse", command=select_input_file)
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
