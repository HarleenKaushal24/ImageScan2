# -*- coding: utf-8 -*-
"""
Created on Tue Mar 11 13:58:23 2025

@author: Harleen
"""

import streamlit as st
import pandas as pd
import numpy as np
import cv2
import os
from PIL import Image
#pip install Office365-REST-Python-Client
from office365.sharepoint.client_context import ClientContext
from io import BytesIO

# Load SharePoint credentials from secrets
SHAREPOINT_SITE = st.secrets["SHAREPOINT_SITE"]
SHAREPOINT_FOLDER = st.secrets["SHAREPOINT_FOLDER"]
TENANT_ID = st.secrets["TENANT_ID"]
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]

# Authenticate with SharePoint
def get_sharepoint_context():
    from office365.runtime.auth.client_credential import ClientCredential
    ctx = ClientContext(f"https://{SHAREPOINT_SITE}").with_credentials(ClientCredential(CLIENT_ID, CLIENT_SECRET))
    return ctx

# Function to fetch file from SharePoint
def fetch_file_from_sharepoint(filename):
    ctx = get_sharepoint_context()
    file_url = f"{SHAREPOINT_FOLDER}/{filename}"
    file = ctx.web.get_file_by_server_relative_url(file_url).download().execute_query()
    return BytesIO(file.content)

# Load Excel file from SharePoint
def load_excel():
    file_data = fetch_file_from_sharepoint("img1.xlsm")
    return pd.read_excel(file_data, engine="openpyxl")

# Load images dynamically from SharePoint
def get_image_files():
    ctx = get_sharepoint_context()
    folder = ctx.web.get_folder_by_server_relative_url(SHAREPOINT_FOLDER)
    files = folder.files.get().execute_query()
    return [file.properties["Name"] for file in files]

# Function to process uploaded image
def process_uploaded_image(uploaded_file):
    image = np.array(Image.open(uploaded_file))
    return cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

# Streamlit UI
def main():
    st.title("Secure Image Recognition App")

    # Check for required files in SharePoint
    try:
        df = load_excel()
        available_images = get_image_files()
    except Exception as e:
        st.error(f"Error fetching files: {e}")
        return

    uploaded_file = st.file_uploader("Upload an image", type=["jpg", "png", "jpeg"])
    
    if uploaded_file:
        target_img = process_uploaded_image(uploaded_file)
        st.image(target_img, caption="Uploaded Image", use_column_width=True)

        if st.button("Find Best Matches"):
            find_best_matches(target_img, df, available_images)

def find_best_matches(target_img, df, available_images):
    """Find matching images stored in SharePoint."""
    ctx = get_sharepoint_context()
    
    sift = cv2.SIFT_create()
    target_kp, target_des = sift.detectAndCompute(cv2.cvtColor(target_img, cv2.COLOR_BGR2GRAY), None)

    match_scores = {}
    for img_name in available_images:
        img_data = fetch_file_from_sharepoint(img_name)
        img = np.array(Image.open(img_data))
        img_gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        
        kp, des = sift.detectAndCompute(img_gray, None)
        if des is None:
            continue

        bf = cv2.BFMatcher(cv2.NORM_L2, crossCheck=True)
        matches = sorted(bf.match(target_des, des), key=lambda x: x.distance)
        
        match_scores[img_name] = len(matches)

    sorted_matches = sorted(match_scores.items(), key=lambda x: x[1], reverse=True)[:2]
    
    st.write("### Top Matches")
    for rank, (img_name, score) in enumerate(sorted_matches, start=1):
        st.write(f"**Rank {rank}: {img_name} ({score} matches)**")

if __name__ == "__main__":
    main()
