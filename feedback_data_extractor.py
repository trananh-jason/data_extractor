# Author: Jason Tran
# Last Update: 24/07/2024
# 
# This is a program created to read data from the Canadian Musicians Cooperative feedback files. As of the last update, this program only works for .xlsx files.
#
# INSTRUCTIONS:
# To use the program, 
# 1. Download the .xlsx feedback file you wish to process.
# 2. Copy the file path of the .xlsx file and paste it on line 63 
#       Line 63 should look like:
#           data_file = ".../your_file.xlsx"
# 3. Press run
# 4. In your terminal the data will be represented in this format
#       {'Experience': {...}, 'Relevancy': {...}, 'Comprehension': {...}, 'Usefulness': {...}}
# 5. The rating tally is represented as a key value pair. For example, {5: 11, 4: 2} means that there are eleven 5 star ratings and two 4 star ratings

import pandas as pd

def calc_average(data: dict) -> float:
    """
    Calculates the average of a specified dictionary

    Parameters:
        data (dict): A dictionary containing the rating tallies
    
    Returns:
        average (float): A weighted average of the rating rallies
    """
    # Count the amount of people who gave a rating
    num_entries = 0
    for i in data.values():
        num_entries += i
    
    sum = 0

    for key, value in data.items():
        sum += key * value
    
    return sum / num_entries

def tally_column(col: str, data_set: 'pandas.core.frame.DataFrame') -> dict:
    """
    Reads the data in a specified column and tallies the entries
    
    Paramaters:
        col (str): A string that represents the header of the column that will be tallied
    
    Returns:
        tally (dict): A dictionary that returns the tally for each unique entry
    """
    tally = {}

    for entry in data_set.get(col):
        if entry in tally.keys():
            tally[entry] += 1
        else:
            tally[entry] = 1
    
    return tally


# data_file contains the file path of the .xlsx file the program will read
data_file = "feedback.xlsx"

# ratings will contain the experience, relevance, comprehension, and usefulness rating data
ratings = {"Experience" : None, "Relevancy" : None, "Comprehension" : None, "Usefulness" : None}

# Extract the columns containing the ratings
df = pd.read_excel(data_file, 
                   usecols=[
                            "How would you rate your experience at the event", 
                            "How would you rate the relevance of the event for your career planning?", 
                            "How would you describe the ease of understanding the session content?", 
                            "How would you rate the usefulness of the information provided to you in this session?"
                        ])

# Function to reduce typing


# Tally columns for their respective sections
ratings["Experience"] = tally_column("How would you rate your experience at the event", df)
ratings["Relevancy"] = tally_column("How would you rate the relevance of the event for your career planning?", df)
ratings["Comprehension"] = tally_column("How would you describe the ease of understanding the session content?", df)
ratings["Usefulness"] = tally_column("How would you rate the usefulness of the information provided to you in this session?", df)

# Calculate rating averages
experience_average = round(calc_average(ratings["Experience"]), 2)
relevancy_average = round(calc_average(ratings["Relevancy"]), 2)
comprehension_average = round(calc_average(ratings["Comprehension"]), 2)
usefulness_average = round(calc_average(ratings["Usefulness"]), 2)

# Show results
print(ratings)
print(f"Average Experience Rating:\t{experience_average}")
print(f"Average Relevancy Rating:\t{relevancy_average}")
print(f"Average Comprehension Rating:\t{comprehension_average}")
print(f"Average Usefulness Rating:\t{usefulness_average}")