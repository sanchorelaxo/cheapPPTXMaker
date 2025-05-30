import os
import datetime
import json
from pptx import Presentation
from midjourney_api import MidjourneyClient  # Hypothetical Midjourney API client

# Load slides from external JSON file
with open("slides.json", "r", encoding="utf-8") as f:
    slides = json.load(f)

# Initialize Midjourney API client (replace with your actual API key/setup)
midjourney_client = MidjourneyClient(api_key="YOUR_API_KEY")

def generate_image(keywords, slide_num):
    """
    Generate an image using the Midjourney API based on the provided keywords.
    Ensure the image is in portrait orientation and PNG format.
    
    Args:
        keywords (str): The keywords to use as the API prompt.
        slide_num (int): Slide number for naming the output file.
    
    Returns:
        str: Path to the generated image, or None if generation fails.
    """
    try:
        # Create a directory for temporary images if it doesn’t exist
        os.makedirs("generated_images", exist_ok=True)
        
        # Set parameters for portrait orientation and PNG format
        response = midjourney_client.generate_image(
            prompt=keywords,
            aspect_ratio="9:16",  # Portrait orientation (adjust as needed)
            animation=False,      # Ensure static image
            output_format="png"   # Request PNG format
        )
        
        # Assume the API returns image data; adjust based on actual API response
        image_path = f"generated_images/slide_{slide_num}.png"
        # Simulate saving the image locally (replace with actual download logic if needed)
        with open(image_path, "wb") as f:
            f.write(response['image_data'])  # Hypothetical response field for image data
        
        return image_path
    except Exception as e:
        print(f"Error generating image for slide {slide_num} using keywords '{keywords}': {e}")
        return None

def create_slide(prs, slide_info, i):
    """
    Create a slide based on provided slide information.
    
    Args:
        prs (Presentation): PowerPoint presentation object.
        slide_info (dict): Dictionary containing slide details.
        i (int): Slide index for naming purposes.
    """
    title = slide_info["title"]
    text = slide_info["text"]
    notes = slide_info["notes"]

    # Use 'Two Content' layout (index 3) for slides with images
    slide_layout = prs.slide_layouts[3]  # Assumes layout 3 is 'Two Content'
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = title

    # Add text to the left content placeholder
    left_placeholder = slide.placeholders[1]
    left_placeholder.text = text

    # Generate and add image to the right content placeholder
    image_path = generate_image(slide_info["keywords"], i)
    if image_path and os.path.exists(image_path):
        right_placeholder = slide.placeholders[2]
        left = right_placeholder.left
        top = right_placeholder.top
        width = right_placeholder.width
        height = right_placeholder.height
        slide.shapes.add_picture(image_path, left, top, width, height)

    # Add presenter's notes
    notes_slide = slide.notes_slide
    notes_text_frame = notes_slide.notes_text_frame
    notes_text_frame.text = notes

def create_presentation(slides, base_filename="presentation"):
    """
    Generate a PowerPoint presentation based on provided slide information.
    
    Args:
        slides (list): List of dictionaries containing slide details (title, text, visual_description, notes).
        base_filename (str): Base name for the output file (default: 'presentation').
    """
    # Initialize a new presentation
    prs = Presentation()

    for i, slide_info in enumerate(slides):
        create_slide(prs, slide_info, i)

    # Generate filename with timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{base_filename}_{timestamp}.pptx"
    prs.save(filename)
    print(f"Presentation saved as {filename}")

    # Clean up generated images
    for i in range(len(slides)):
        image_path = f"generated_images/slide_{i}.png"
        if os.path.exists(image_path):
            os.remove(image_path)
    if os.path.exists("generated_images") and not os.listdir("generated_images"):
        os.rmdir("generated_images")

if __name__ == "__main__":
    create_presentation(slides)