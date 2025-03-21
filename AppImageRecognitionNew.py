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

# Preprocessing function to standardize images
def preprocess_image(img):
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)  # Convert to grayscale
    img_resized = cv2.resize(img_gray, (300, 300))  # Resize image for consistency
    img_equalized = cv2.equalizeHist(img_resized)  # Apply histogram equalization
    img_denoised = cv2.GaussianBlur(img_equalized, (5, 5), 0)  # Apply Gaussian blur to reduce noise
    return img_denoised

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

    drive_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_site}/drives"
    drive_response = requests.get(drive_url, headers=headers)
    drive_id = drive_response.json()["value"][0]["id"]

    folder_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/ImageFiles/extracted_images:/children"
    response = requests.get(folder_url, headers=headers)

    if response.status_code == 200:
        images = response.json()["value"]
        image_urls = [img["@microsoft.graph.downloadUrl"] for img in images]

        with concurrent.futures.ThreadPoolExecutor() as executor:
            img_responses = list(executor.map(requests.get, image_urls))

        return [(img["name"], response) for img, response in zip(images, img_responses)]
    else:
        st.error("Failed to fetch images from SharePoint folder")
        return []

# Function to process the fetched images
def process_fetched_images():
    img_responses = fetch_images_from_sharepoint_parallel()
    if img_responses:
        images = []
        for filename, response in img_responses:
            if response.status_code == 200:
                img_color = np.array(Image.open(BytesIO(response.content)))
                img_color_preprocessed = preprocess_image(img_color)  # Preprocess fetched image
                images.append((filename, img_color_preprocessed))
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
    img.save("barcode.png", format="PNG")  # Save the barcode image as PNG to ensure lossless quality
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
        target_img = np.array(Image.open(uploaded_file))
        target_img = cv2.cvtColor(target_img, cv2.COLOR_RGB2BGR)

        # Call the function to find best matches
        find_best_matches(target_img, df)

def find_best_matches(target_img, df):
    """Find and display the best matching image from SharePoint."""
    try:
        with st.spinner("Finding best match..."):
            target_img_gray = preprocess_image(target_img)  # Preprocess the target image

            sift = cv2.SIFT_create()
            target_kp, target_des = sift.detectAndCompute(target_img_gray, None)

            if target_des is None:
                st.error("No features detected in the uploaded image.")
                return

            sharepoint_images = process_fetched_images()

            match_scores = {}

            # Loop through SharePoint images and match them
            for filename, img_gray in sharepoint_images:
                kp, des = sift.detectAndCompute(img_gray, None)
                if des is None:
                    continue

                # Feature matching using FLANN-based matcher
                FLANN_INDEX_KDTREE = 1
                index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
                search_params = dict(checks=50)

                flann = cv2.FlannBasedMatcher(index_params, search_params)
                matches = flann.knnMatch(target_des, des, k=2)

                # Apply Lowe's ratio test to filter out poor matches
                good_matches = []
                for m, n in matches:
                    if m.distance < 0.7 * n.distance:
                        good_matches.append(m)

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
    img_name, match_count = top_match

    st.write(f"**Top Match: {img_name} ({match_count} good matches)**")

    matched_row = df[df['ImageName'] == img_name]
    if matched_row is not None:
        itm_value = matched_row["ITM"].values[0]
        BC = str(matched_row["Barcode"].values[0])
        st.write(f"**ITM Name**: {itm_value}")
        
        barcode_img = generate_barcode_image(BC)
        st.image(barcode_img, caption=f"Barcode for {img_name}", use_container_width=True)
        
    else:
        st.error("No matching data found in Excel.")

if __name__ == "__main__":
    main()
