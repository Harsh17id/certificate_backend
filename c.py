import pandas as pd
import os
import streamlit as st
from streamlit_image_coordinates import streamlit_image_coordinates
from PIL import Image, ImageDraw, ImageFont
import random
import string


def generate_code():
    """Generate a random 8-character alphanumeric code."""
    chars = string.ascii_uppercase + string.digits
    return ''.join(random.choices(chars, k=8))


def read_excel(uploaded_file, sheet_name=None):
    """Read an Excel file and return a list of names."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        if isinstance(df, dict):  # If multiple sheets, pick first one
            df = df[sheet_name] if sheet_name else list(df.values())[0]
        return df.iloc[:, 0].dropna().astype(str).tolist()
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None


def load_template(uploaded_template):
    try:
        return Image.open(uploaded_template).convert("RGB")
    except Exception as e:
        st.error(f"Error loading template: {e}")
        return None


def add_name_to_certificate(template, name, position, font_size=50):
    cert = template.copy()
    draw = ImageDraw.Draw(cert)

    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except IOError:
        st.warning("Font 'arial.ttf' not found, using default font.")
        font = ImageFont.load_default()

    # Get text bounding box
    bbox = draw.textbbox((0, 0), name, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    # Adjust position to center text
    text_position = (position[0] - text_width // 2, position[1] - text_height // 2)

    # Draw text
    draw.text(text_position, name, fill=(0, 0, 0), font=font)
    return cert


def save_certificate(image, name):
    """Save the generated certificate in a structured folder."""
    os.makedirs(f'name', exist_ok=True)
    filename = os.path.join(f'name', f"{generate_code()}.png")
    image.save(filename)
    return filename


# Streamlit UI
st.title("Certificate Generator")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx", "ods"])
uploaded_template = st.file_uploader("Upload Certificate Template (JPG/PNG)", type=["jpg", "png"])
font_size = st.number_input("Enter font size", min_value=10, max_value=100, value=50)

if uploaded_file and uploaded_template:
    names = read_excel(uploaded_file)
    template = load_template(uploaded_template)

    if names and template:
        # Resize image for easier selection (maintain aspect ratio)
        preview_width = 600  # Adjust preview size
        aspect_ratio = template.height / template.width
        preview_height = int(preview_width * aspect_ratio)
        preview_template = template.resize((preview_width, preview_height))

        st.image(preview_template, caption="Certificate Template Preview", use_container_width=False)
        coords = streamlit_image_coordinates(preview_template)

        if coords:
            if "x" in coords and "y" in coords:
                st.write(f"Selected position: {coords}")
                scale_x = template.width / preview_width
                scale_y = template.height / preview_height
                x, y = int(coords["x"] * scale_x), int(coords["y"] * scale_y)

                for name in names:
                    cert = add_name_to_certificate(template, name, (x, y), font_size)
                    file_path = save_certificate(cert, name)
                    st.image(cert, caption=f"Generated Certificate for {name}")
                    st.write(f"Saved: {file_path}")

                st.success("All certificates generated successfully!")
            else:
                st.warning("Invalid selection! Please click on the image to select a position.")
        else:
            st.warning("Click on the image to select a text position.")
