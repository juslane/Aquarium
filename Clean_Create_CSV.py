# Import libraries

import pandas as pd
from tkinter import filedialog
from tkinter import *
import random

# Set up root for tkinter

root = Tk()

# Ask user to select the file and create the dataframe to clean

root.fileName = filedialog.askopenfilename(filetypes = (("Excel Files", "*.xlsx"), ("All File Types", "*.*")))
xl_messy = pd.read_excel(root.fileName, header = 1)

# Select the 'Email Address' and 'First Name' columns

df_messy_Email = pd.DataFrame(xl_messy['Email Address'])  
df_messy_Name = pd.DataFrame(xl_messy['First Name']) 

# Put the columns together in the correct order

df_tidy = pd.concat([df_messy_Email, df_messy_Name], axis = 1)

# Rename 'Email Address' column to 'Email'

df_tidy.columns = ['Email', 'First Name']

# Clean further by eliminating rows that do not have an '@' and do not have a '.' in the 'Email' column; show them to the user

print("Here is the list of email addresses that failed to pass the test:")

bad_emails = df_tidy[(df_tidy.Email.str.contains("@") == False) & (df_tidy.Email.str.contains("\.") == False)]

indices = []

for idx, email in enumerate(bad_emails['Email']):
    print(idx, email, bad_emails['First Name'][idx])
    indices.append(idx)

# Eliminate the necessary rows

tidy_emails_na = df_tidy.drop(df_tidy[(df_tidy.Email.str.contains("@") == False) & (df_tidy.Email.str.contains("\.") == False)].index)

tidy_emails = tidy_emails_na[pd.notnull(tidy_emails_na['Email'])]

# Reset the indices

indexed_emails = tidy_emails.reset_index(drop = True)

# Randomly assign number to each visitor, then sort.

random_assignment = pd.Series(random.sample(range(1, len(indexed_emails) + 1), len(indexed_emails)))

indexed_emails.index = [random_assignment]

randomized_emails = indexed_emails.sort_index()

# Write half of the Dataframe to Scale.csv and the other half to OER.csv

scale_emails = randomized_emails[round(len(randomized_emails) / 2):]
oer_emails = randomized_emails[:round(len(randomized_emails) / 2)]

export_file_path_oer = filedialog.asksaveasfilename(defaultextension='.csv')
oer_emails.to_csv(export_file_path_oer, index = False)

export_file_path_scale = filedialog.asksaveasfilename(defaultextension='.csv')
scale_emails.to_csv(export_file_path_scale, index = False)

# Tell the user how many emails are in each list

num_scale = len(scale_emails)
num_oer = len(oer_emails)

print("There are " + str(num_scale) + " Scale emails, and " + str(num_oer) + " OER emails ready to be cleaned.")
