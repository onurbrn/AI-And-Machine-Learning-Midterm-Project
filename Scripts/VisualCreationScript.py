import openai
import os
import requests
from docx import Document

# Set your OpenAI API key
openai.api_key = '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'

def read_docx(file_path):
    """
    Read the content of a .docx file.
    
    Args:
        file_path (str): Path to the .docx file
        
    Returns:
        str: The combined text content of the document
    """
    doc = Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

def generate_visualization(prompt):
    """
    Generate an image based on the given prompt using OpenAI's DALL-E 3 model.
    
    Args:
        prompt (str): The description of the image to generate
        
    Returns:
        str: URL of the generated image
    """
    response = openai.Image.create(
        prompt=prompt,
        n=1,  # Generate one image
        size="1024x1024"  # Image size
    )
    return response['data'][0]['url']

def download_image(image_url, save_path):
    """
    Download an image from a URL and save it to a file.
    
    Args:
        image_url (str): The URL of the image
        save_path (str): The path where the image will be saved
    """
    response = requests.get(image_url, stream=True)
    if response.status_code == 200:
        with open(save_path, 'wb') as file:
            for chunk in response.iter_content(1024):
                file.write(chunk)
    else:
        print(f"Failed to download image from {image_url}")

def process_visual_descriptions(folder_path, output_folder):
    """
    Generate crime scene visuals based on documents in the Visual Descriptions folder.
    
    Args:
        folder_path (str): Path to the folder containing .docx files
        output_folder (str): Path to the folder where images will be saved
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".docx"):
            file_path = os.path.join(folder_path, file_name)
            content = read_docx(file_path)
            
            # Generate visualization for the document content
            visual_prompt = f"Create a highly detailed and artistic visualization of the following crime scene description: {content[:500]}..."
            image_url = generate_visualization(visual_prompt)
            
            # Save the image
            image_name = f"crime_scene_{os.path.splitext(file_name)[0]}.png"
            save_path = os.path.join(output_folder, image_name)
            download_image(image_url, save_path)
            print(f"Saved crime scene visualization for {file_name} as {image_name}")

def process_other_documents(folder_path, output_folder):
    """
    Generate creative visuals for evidence and objects from other documents.
    
    Args:
        folder_path (str): Path to the folder containing other .docx files
        output_folder (str): Path to the folder where images will be saved
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".docx"):
            file_path = os.path.join(folder_path, file_name)
            content = read_docx(file_path)
            
            # Generate visualization for evidence or key objects
            visual_prompt = f"Create a realistic and creative representation of evidence or key objects from the following text: {content[:500]}..."
            image_url = generate_visualization(visual_prompt)
            
            # Save the image
            image_name = f"evidence_{os.path.splitext(file_name)[0]}.png"
            save_path = os.path.join(output_folder, image_name)
            download_image(image_url, save_path)
            print(f"Saved evidence visualization for {file_name} as {image_name}")

def create_additional_visuals(output_folder):
    """
    Generate additional creative visuals using unique prompts unrelated to the documents.
    
    Args:
        output_folder (str): Path to the folder where images will be saved
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    prompts = [
        "A futuristic crime lab with holographic evidence displays and advanced forensic tools.",
        "A blood-stained weapon lying in a dimly lit alleyway, surrounded by police markers.",
        "A courtroom scene with lawyers, a judge, and visual evidence displayed on a screen."
    ]

    for i, prompt in enumerate(prompts, start=1):
        image_url = generate_visualization(prompt)
        
        # Save the image
        image_name = f"additional_visual_{i}.png"
        save_path = os.path.join(output_folder, image_name)
        download_image(image_url, save_path)
        print(f"Saved additional visualization {i} as {image_name}")

# Example Usage:
visual_descriptions_folder = "C:\\Users\\copro\\OneDrive\\Masaüstü\\machinepicture"  # Folder containing crime scene descriptions
other_documents_folder = "C:\\Users\\copro\\OneDrive\\Masaüstü\\machinepicture"         # Folder containing evidence-related or other .docx files
output_folder_crime_scene = "C:\\Users\\copro\\OneDrive\\Masaüstü\\machinepicture2\\Crime_Scene_Visuals"   # Output folder for crime scene visuals
output_folder_evidence = "C:\\Users\\copro\\OneDrive\\Masaüstü\\machinepicture2\\Evidence_Visuals"        # Output folder for evidence visuals
output_folder_additional = "C:\\Users\\copro\\OneDrive\\Masaüstü\\machinepicture2\\Additional_Visuals"    # Output folder for additional visuals

# Generate visuals
process_visual_descriptions(visual_descriptions_folder, output_folder_crime_scene)
process_other_documents(other_documents_folder, output_folder_evidence)
create_additional_visuals(output_folder_additional)