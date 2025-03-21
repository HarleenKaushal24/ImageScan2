# -*- coding: utf-8 -*-
"""
Created on Fri Mar 21 12:48:15 2025

@author: Harleen
"""

import streamlit as st
import pandas as pd
import cv2
import numpy as np
import requests
from io import BytesIO
from PIL import Image
from matplotlib import pyplot as plt
import barcode
from barcode.writer import ImageWriter
import concurrent.futures

# Constants
TENANT_ID = st.secrets["TENANT_ID"]
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
SHAREPOINT_SITE = "boscoandroxysinc.sharepoint.com:/sites/BnR-Data"
FILE_PATH = "/ImageFiles/img1.xlsm"
IMAGE_FOLDER = "extracted_images"

# Function to get access token
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = requests.post(url, data=payload, headers=headers)
    return response.json().get("access_token")

# Function to fetch Excel file from SharePoint
def fetch_excel_from_sharepoint():
    access_token = get_access_token()
    if not access_token:
        st.error("Failed to authenticate with SharePoint")
        return None
    
    headers = {"Authorization": f"Bearer {access_token}"}
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}"
    site_response = requests.get(site_url, headers=headers)
    sharepoint_site = site_response.json()["id"].split(",")[1]
    
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_site}/drives"
    drive_response = requests.get(drive_url, headers=headers)
    drive_id = drive_response.json()["value"][0]["id"]
    
    file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{FILE_PATH}:/content"
    response = requests.get(file_url, headers=headers)
    
    if response.status_code == 200:
        return BytesIO(response.content)
    else:
        st.error("Failed to fetch file from SharePoint")
        return None

# Function to fetch images from SharePoint in parallel
def fetch_images_from_sharepoint_parallel():
    access_token = get_access_token()
    if not access_token:
        st.error("Failed to authenticate with SharePoint")
        return []

    headers = {"Authorization": f"Bearer {access_token}"}
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}"
    site_response = requests.get(site_url, headers=headers)
    sharepoint_site = site_response.json()["id"].split(",")[1]

    # Get the drive ID (this part is shared by both functions)
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_site}/drives"
    drive_response = requests.get(drive_url, headers=headers)
    drive_id = drive_response.json()["value"][0]["id"]

    # Fetch the list of image files in the folder
    folder_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/ImageFiles/extracted_images:/children"
    response = requests.get(folder_url, headers=headers)

    if response.status_code == 200:
        # Extract image URLs
        images = response.json()["value"]
        image_urls = [img["@microsoft.graph.downloadUrl"] for img in images]

        # Use ThreadPoolExecutor to fetch images concurrently
        with concurrent.futures.ThreadPoolExecutor() as executor:
            img_responses = list(executor.map(requests.get, image_urls))

        # Return the images as responses along with filenames
        return [(img["name"], response) for img, response in zip(images, img_responses)]
    else:
        st.error("Failed to fetch images from SharePoint folder")
        return []

# Function to process the fetched images
def process_fetched_images():
    # Fetch the images concurrently
    img_responses = fetch_images_from_sharepoint_parallel()

    # Process each image (you can further handle the responses here)
    if img_responses:
        images = []
        for filename, response in img_responses:
            if response.status_code == 200:
                img_color = np.array(Image.open(BytesIO(response.content)))
                images.append((filename, img_color))
        return images
    else:
        st.error("No images fetched or error occurred.")
        return []

# Function to generate barcode as an image
def generate_barcode_image(barcode_string):
    code128 = barcode.get_barcode_class('code128')  # Use Code128
    barcode_instance = code128(barcode_string, writer=ImageWriter())

    options = {
        'module_width': 0.6,  # Increase thickness of bars
        'module_height': 25,  # Increase height of the barcode
        'font_size': 18,  # Adjust the size of the numbers beneath the barcode
        'text_distance': 10,  # Distance between the barcode and the text
    }

    img_stream = BytesIO()
    barcode_instance.write(img_stream, options=options)
    img_stream.seek(0)

    img = Image.open(img_stream)

    # Save the barcode image as PNG to ensure lossless quality
    img.save("barcode.png", format="PNG")
    return img

# Main function to run the Streamlit app
def main():
    """Main Streamlit App."""
    st.title("Test APP")

    # Fetch the Excel file from SharePoint
    excel_file = fetch_excel_from_sharepoint()
    if not excel_file:
        return

    # Read the Excel data into a dataframe
    df = pd.read_excel(excel_file, engine='openpyxl')

    # Upload the image (no display of the uploaded image)
    uploaded_file = st.file_uploader("Upload an image", type=["jpg", "png", "jpeg"])

    if uploaded_file:
        # Once image is uploaded, start the matching process
        target_img = np.array(Image.open(uploaded_file))
        target_img = cv2.cvtColor(target_img, cv2.COLOR_RGB2BGR)

        # Call the function to find best matches
        find_best_matches(target_img, df)

def find_best_matches(target_img, df):
    """Find and display the best matching image from SharePoint."""
    try:
        # Show progress bar while processing
        with st.spinner("Finding best match..."):
            # Convert target image to grayscale and resize for matching
            target_img_resized = cv2.resize(target_img, (300, 300))
            target_img_gray = cv2.cvtColor(target_img_resized, cv2.COLOR_BGR2GRAY)

            # Apply histogram equalization for better contrast
            target_img_gray = cv2.equalizeHist(target_img_gray)

            # Optional: Apply Gaussian Blur to reduce noise
            target_img_gray = cv2.GaussianBlur(target_img_gray, (5, 5), 0)

            # SIFT feature detection
            sift = cv2.SIFT_create()
            target_kp, target_des = sift.detectAndCompute(target_img_gray, None)

            if target_des is None:
                st.error("No features detected in the uploaded image.")
                return

            # Fetch images from SharePoint
            sharepoint_images = process_fetched_images()

            match_scores = {}

            # Loop through SharePoint images and match them
            for filename, img_color in sharepoint_images:
                # Convert to grayscale and remove background
                img_gray = cv2.cvtColor(remove_background(img_color), cv2.COLOR_BGR2GRAY)

                # Apply histogram equalization to the SharePoint image
                img_gray = cv2.equalizeHist(img_gray)

                # Optional: Apply Gaussian Blur
                img_gray = cv2.GaussianBlur(img_gray, (5, 5), 0)

                # Detect keypoints and descriptors in the SharePoint image
                kp, des = sift.detectAndCompute(img_gray, None)
                if des is None:
                    continue

                # Feature matching using FLANN-based matcher
                FLANN_INDEX_KDTREE = 1
                index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
                search_params = dict(checks=50)  # Higher number of checks for better accuracy

                flann = cv2.FlannBasedMatcher(index_params, search_params)
                matches = flann.knnMatch(target_des, des, k=2)

                # Apply Lowe's ratio test to filter out poor matches
                good_matches = []
                for m, n in matches:
                    if m.distance < 0.7 * n.distance:
                        good_matches.append(m)

                # Store match information using the image filename
                match_scores[filename] = len(good_matches)

            # Sort the matches by the number of good matches, and get the top match
            top_match = sorted(match_scores.items(), key=lambda x: x[1], reverse=True)[0]

            # Display the top match and ITM value
            display_results(top_match, target_img_gray, target_kp, df)

    except Exception as e:
        st.error(f"Unexpected error: {e}")

# Function to display the results
def display_results(top_match, target_img_gray, target_kp, df):
    """Display the top matching image and ITM value along with barcode."""
    # Extract match data
    img_name, match_count = top_match

    # Display match count and ITM value from Excel
    st.write(f"**Top Match: {img_name} ({match_count} good matches)**")

    # Extract row number from filename (assuming filename format is consistent)
    matched_row = df[df['ImageName'] == img_name]
    if matched_row is not None:
        itm_value = matched_row["ITM"].values[0]
        BC = str(matched_row["Barcode"].values[0])
        st.write(f"**ITM Name**: {itm_value}")
        
        # Generate and display barcode image for the matched ITM
        barcode_img = generate_barcode_image(BC)
        st.image(barcode_img, caption=f"Barcode for {img_name}", use_container_width=True)
        
    else:
        st.error("No matching data found in Excel.")

def remove_background(img):
    """Remove white background."""
    try:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if contours:
            x, y, w, h = cv2.boundingRect(contours[0])
            img = img[y:y+h, x:x+w]
        return img
    except Exception:
        return img

if __name__ == "__main__":
    main()
