# -*- coding: utf-8 -*-
"""
Created on Tue Mar 11 13:58:23 2025

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



def fetch_images_from_sharepoint():
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
        return response.json()["value"]
    else:
        st.error("Failed to fetch images from SharePoint folder")
        return []


def main():
    """Main Streamlit App."""
    st.title("Image Recognition and Matching")
    
    # Fetch the Excel file from SharePoint
    excel_file = fetch_excel_from_sharepoint()
    if not excel_file:
        return
    
    # Read the Excel data into a dataframe
    df = pd.read_excel(excel_file, engine='openpyxl')

    uploaded_file = st.file_uploader("Upload an image", type=["jpg", "png", "jpeg"])
    
    if uploaded_file:
        target_img = np.array(Image.open(uploaded_file))
        target_img = cv2.cvtColor(target_img, cv2.COLOR_RGB2BGR)
        st.image(target_img, caption="Uploaded Image", use_column_width=True)
        
        if st.button("Find Best Matches"):
            find_best_matches(target_img, df)


def find_best_matches(target_img, df):
    """Find and display best matching images from SharePoint."""
    try:
        # Convert target image to grayscale and resize for matching
        target_img_gray = cv2.cvtColor(cv2.resize(target_img, (300, 300)), cv2.COLOR_BGR2GRAY)

        # SIFT feature detection
        sift = cv2.SIFT_create()
        target_kp, target_des = sift.detectAndCompute(target_img_gray, None)
        
        if target_des is None:
            st.error("No features detected in the uploaded image.")
            return

        # Fetch images from SharePoint
        sharepoint_images = fetch_images_from_sharepoint()

        match_scores = {}
        
        for image_info in sharepoint_images:
            # Get image URL from SharePoint
            image_url = image_info["@microsoft.graph.downloadUrl"]
            img_response = requests.get(image_url)
            img_color = np.array(Image.open(BytesIO(img_response.content)))

            # Process and match
            img_gray = cv2.cvtColor(remove_background(img_color), cv2.COLOR_BGR2GRAY)
            kp, des = sift.detectAndCompute(img_gray, None)
            if des is None:
                continue

            # Feature matching
            bf = cv2.BFMatcher(cv2.NORM_L2, crossCheck=True)
            matches = sorted(bf.match(target_des, des), key=lambda x: x.distance)

            match_scores[image_info["name"]] = (len(matches), img_gray, img_color, kp, matches)

        # Sort the matches by number of good matches
        top_matches = sorted(match_scores.items(), key=lambda x: x[1][0], reverse=True)[:2]
        
        display_results(top_matches, target_img_gray, target_kp, df)
    
    except Exception as e:
        st.error(f"Unexpected error: {e}")


def display_results(top_matches, target_img_gray, target_kp, df):
    """Display top matching images and save results."""
    if not top_matches:
        st.write("No matches found.")
        return

    st.write("### Top 2 Best Matches")
    match_results = []

    for rank, (img_name, (match_count, img_gray, img_color, kp, matches)) in enumerate(top_matches, start=1):
        st.write(f"**Rank {rank}: {img_name} ({match_count} good matches)**")

        matched_row = extract_row_number(img_name, len(df))
        if matched_row is not None:
            matched_data = df.iloc[matched_row]
            st.write(matched_data)

            match_results.append({
                "Rank": rank,
                "Image Name": img_name,
                "Good Matches": match_count,
                "Excel Row": matched_row,
                **matched_data.to_dict()
            })

        # Display matched images
        fig, ax = plt.subplots(1, 2, figsize=(10, 5))
        ax[0].imshow(cv2.drawKeypoints(target_img_gray, target_kp, None, color=(0, 255, 0)), cmap='gray')
        ax[0].set_title("Target Image")
        ax[1].imshow(cv2.drawKeypoints(img_gray, kp, None, color=(255, 0, 0)), cmap='gray')
        ax[1].set_title(f"Matched Image {rank}")
        st.pyplot(fig)

    # Save results
    #pd.DataFrame(match_results).to_excel("match_results.xlsx", index=False)
    #st.success("Match results saved to match_results.xlsx")


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


def extract_row_number(filename, df_length):
    """Extract row number safely."""
    try:
        parts = filename.split("_")
        if len(parts) > 1:
            row_number = int(parts[1])
            if 0 <= row_number < df_length:
                return row_number
    except ValueError:
        pass
    return None


if __name__ == "__main__":
    main()