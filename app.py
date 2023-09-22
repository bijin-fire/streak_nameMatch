# import pandas as pd
# import streamlit as st
# from fuzzywuzzy import fuzz
# import io

# # Function to determine test-taker status and match percentage
# def determine_test_taker_status(name, number, df_test_takers, name_threshold, number_threshold):
#     name = name.lower()
    
#     # Initialize match percentages
#     name_match_percentage = 0
#     number_match_percentage = 0
    
#     # Check for fuzzy name matches in 'df_test_takers'
#     for test_name in df_test_takers['Full Name']:
#         name_similarity = fuzz.partial_ratio(name, test_name.lower())
#         if name_similarity >= name_threshold:
#             name_match_percentage = max(name_match_percentage, name_similarity)
    
#     # Check for fuzzy number matches in 'df_test_takers'
#     for test_number in df_test_takers['Mobile Number']:
#         number_similarity = fuzz.partial_ratio(str(number), str(test_number))
#         if number_similarity >= number_threshold:
#             number_match_percentage = max(number_match_percentage, number_similarity)
    
#     # Determine test-taker status based on name and number matches
#     if name_match_percentage > number_match_percentage:
#         test_taker_status = "☑ name_match"
#         match_percentage = name_match_percentage
#     elif number_match_percentage > 0:
#         test_taker_status = "☑ number_match"
#         match_percentage = number_match_percentage
#     else:
#         test_taker_status = " "
#         match_percentage = 0  # Set match percentage to 0 for non-matches
    
#     return test_taker_status, match_percentage

# # Streamlit app header
# st.title("Test-Taker Status Determination")

# # Sidebar for user input
# st.sidebar.header("Settings")
# name_threshold = st.sidebar.slider("Name Similarity Threshold", min_value=0, max_value=100, value=88, step=1)
# number_threshold = st.sidebar.slider("Number Similarity Threshold", min_value=0, max_value=100, value=90, step=1)

# # File upload for 'nfo 6-8' Excel file
# st.sidebar.write("### Upload 'NFO Sheet Data' Excel File")
# uploaded_nfo_file = st.sidebar.file_uploader("Choose a file", type=["xlsx"], key="nfo_file")

# # File upload for 'all_student_names' Excel files
# st.sidebar.write("### Upload 'Test Details' Excel Files")
# uploaded_test_names_files = st.sidebar.file_uploader("Choose files", type=["xlsx"], accept_multiple_files=True, key="test_names_files")

# # Check if files are uploaded
# if uploaded_nfo_file is not None and uploaded_test_names_files is not None:
#     try:
#         # Load the uploaded 'nfo 6-8' Excel file containing all student names and numbers
#         df_all_names = pd.read_excel(uploaded_nfo_file)

#         # Combine multiple uploaded 'all_student_names' Excel sheets into one DataFrame
#         dfs_test_takers = []
#         for uploaded_file in uploaded_test_names_files:
#             df_test_takers = pd.read_excel(uploaded_file)
#             dfs_test_takers.append(df_test_takers)

#         df_test_takers_combined = pd.concat(dfs_test_takers, ignore_index=True)

#         # Display the uploaded data
#         st.write("### Input Data")
#         st.write(df_all_names.head())

#         # Calculate test-taker status and match percentage
#         test_taker_status_list = []
#         match_percentage_list = []

#         for index, row in df_all_names.iterrows():
#             name = row['Name of Student']
#             number = row['Student Number']

#             test_taker_status, match_percentage = determine_test_taker_status(name, number, df_test_takers_combined, name_threshold, number_threshold)

#             test_taker_status_list.append(test_taker_status)
#             match_percentage_list.append(match_percentage)

#         # Add the results to the DataFrame
#         df_all_names['Test_Taker_Status'] = test_taker_status_list
#         df_all_names['Match_Percentage'] = match_percentage_list

#         # Display the results
#         st.write("### Results")
#         st.write(df_all_names)

#         # Export the results to an Excel file
#         if st.button("Export Results to Excel"):
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
#                 df_all_names.to_excel(writer, sheet_name='Sheet1', index=False)
#                 writer.save()
#             output.seek(0)
#             st.download_button(label="Download Results", data=output, file_name='all_names_with_test_taker_status.xlsx', key="download_button")

#     except Exception as e:
#         st.error(f"An error occurred: {e}")

# # About section
# st.sidebar.markdown("#### About")
# st.sidebar.info(
#     "This app calculates test-taker status and match percentage based on fuzzy matching "
#     "for both name and number columns. You can adjust the similarity thresholds using the sidebar."
# )

import pandas as pd
import streamlit as st
from fuzzywuzzy import fuzz
import io

# Function to determine test-taker status and match percentage
def determine_test_taker_status(name, number, df_test_takers, name_threshold, number_threshold):
    name = name.lower()
    
    # Initialize match percentages
    name_match_percentage = 0
    number_match_percentage = 0
    
    # Check for fuzzy name matches in 'df_test_takers'
    for test_name in df_test_takers['Full Name']:
        name_similarity = fuzz.partial_ratio(name, test_name.lower())
        if name_similarity >= name_threshold:
            name_match_percentage = max(name_match_percentage, name_similarity)
    
    # Check for fuzzy number matches in 'df_test_takers'
    for test_number in df_test_takers['Mobile Number']:
        number_similarity = fuzz.partial_ratio(str(number), str(test_number))
        if number_similarity >= number_threshold:
            number_match_percentage = max(number_match_percentage, number_similarity)
    
    # Determine test-taker status based on name and number matches
    if name_match_percentage > number_match_percentage:
        test_taker_status = "☑ name_match"
        match_percentage = name_match_percentage
    elif number_match_percentage > 0:
        test_taker_status = "☑ number_match"
        match_percentage = number_match_percentage
    else:
        test_taker_status = " "
        match_percentage = 0  # Set match percentage to 0 for non-matches
    
    return test_taker_status, match_percentage

# Streamlit app header
st.title("Test-Taker Status Determination")

# Sidebar for user input
st.sidebar.header("Settings")
name_threshold = st.sidebar.slider("Name Similarity Threshold", min_value=0, max_value=100, value=88, step=1)
number_threshold = st.sidebar.slider("Number Similarity Threshold", min_value=0, max_value=100, value=90, step=1)

# File upload for 'nfo 6-8' Excel file
st.sidebar.write("### Upload 'NFO Sheet Data' Excel File")
uploaded_nfo_file = st.sidebar.file_uploader("Choose a file", type=["xlsx"], key="nfo_file")

# File upload for 'all_student_names' Excel files
st.sidebar.write("### Upload 'Test Details' Excel Files")
uploaded_test_names_files = st.sidebar.file_uploader("Choose files", type=["xlsx"], accept_multiple_files=True, key="test_names_files")

# Check if files are uploaded
if uploaded_nfo_file is not None and uploaded_test_names_files is not None:
    try:
        # Load the uploaded 'nfo 6-8' Excel file containing all student names and numbers
        df_all_names = pd.read_excel(uploaded_nfo_file)

        # Combine multiple uploaded 'all_student_names' Excel sheets into one DataFrame
        dfs_test_takers = []
        for uploaded_file in uploaded_test_names_files:
            df_test_takers = pd.read_excel(uploaded_file)
            dfs_test_takers.append(df_test_takers)

        df_test_takers_combined = pd.concat(dfs_test_takers, ignore_index=True)

        # Display the uploaded data
        st.write("### Input Data")
        st.write(df_all_names.head())

        # Calculate test-taker status and match percentage
        test_taker_status_list = []
        match_percentage_list = []

        # Create a list to store students who didn't match with NFO Sheet data
        students_not_in_nfo = []

        for index, row in df_all_names.iterrows():
            name = row['Name of Student']
            number = row['Student Number']

            test_taker_status, match_percentage = determine_test_taker_status(name, number, df_test_takers_combined, name_threshold, number_threshold)

            test_taker_status_list.append(test_taker_status)
            match_percentage_list.append(match_percentage)

            # If there is no match, add the student to the list
            if test_taker_status == " ":
                students_not_in_nfo.append(row.to_dict())

        # Add the results to the DataFrame
        df_all_names['Test_Taker_Status'] = test_taker_status_list
        df_all_names['Match_Percentage'] = match_percentage_list

        # Display the results
        st.write("### Results")
        st.write(df_all_names)

        # Create a DataFrame for students not in NFO Sheet
        df_students_not_in_nfo = pd.DataFrame(students_not_in_nfo)

        # Display students not in NFO Sheet
        st.write("### Students Not in NFO Sheet")
        st.write(df_students_not_in_nfo)

        # Export the results and students not in NFO Sheet to Excel files
        if st.button("Export Results to Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_all_names.to_excel(writer, sheet_name='Sheet1', index=False)
                df_students_not_in_nfo.to_excel(writer, sheet_name='Students_Not_in_NFO', index=False)
                writer.save()
            output.seek(0)
            st.download_button(label="Download Results", data=output, file_name='all_names_with_test_taker_status.xlsx', key="download_button")

    except Exception as e:
        st.error(f"An error occurred: {e}")

# About section
st.sidebar.markdown("#### About")
st.sidebar.info(
    "This app calculates test-taker status and match percentage based on fuzzy matching "
    "for both name and number columns. You can adjust the similarity thresholds using the sidebar."
)

