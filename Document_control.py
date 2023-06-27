# it is going to be the edited version 1
import numpy as np
import pandas as pd
from datetime import datetime

# Load data
reg = pd.read_excel("Document_register.xlsx")

# Use second row as column names
reg.columns = reg.iloc[1]

# Drop documents with status "canceled"
reg = reg[reg["Document Status"] != "canceled"]

# Get current date
today = datetime.now()

def outlist(reg):
    """
    Filters the register for documents that need to be processed and
    sorts by submission date.
    """

    # Drop documents that have a TM history
    reg = reg[reg["TM History TM(Rev)"].isna()]

    # Clean up data and sort
    reg = reg.drop(0)
    reg = reg.dropna(subset = ["Clough Doc No"])
    reg["Forecast Submission Date2"] = pd.to_datetime(reg["Forecast Submission Date2"], errors="coerce")
    reg = reg.sort_values(by = "Forecast Submission Date2")

    return reg

def overdue(reg):
    """
    Prints a list of documents that are overdue.
    """

    reg = outlist(reg)

    due_date = reg["Forecast Submission Date2"]
    due_date = due_date.dropna()

    # Calculate days overdue
    overdue = (due_date - today).dt.days

    overduefinal = overdue[overdue < 0]

    print(f"\n\nThere are {len(overduefinal)} overdue documents as of today ({today.date()}) as listed below:\n\n")

    for i, days in overduefinal.iteritems():
        print(f"{i+1} - {reg.loc[i, 'Title']} ({reg.loc[i, 'Clough Doc No']})\n{days} days overdue ({due_date[i].date()})\n")

def outstanding(reg):
    """
    Prints a list of documents that are outstanding.
    """

    reg = outlist(reg)

    print(f"\n\nAnd there are {reg.shape[0]} outstanding documents as listed below:\n\n")

    for i, row in reg.iterrows():
        due_date = row["Forecast Submission Date2"]
        od = (due_date - today).days

        if od > 0:
            print(f"{i+1} - {row['Title']} ({row['Clough Doc No']}) Due in {od} days\n")
        else:
            print(f"{i+1} - {row['Title']} ({row['Clough Doc No']}) overdue {-od} days\n")

def report(reg):
    """
    Prints a report of the submission status of documents.
    """

    regout = outlist(reg)

    # Count submitted documents
    submitted = reg["Forecast Submission Date2"].count() - regout.shape[0]

    print(f"We have submitted {submitted} out of {reg['Forecast Submission Date2'].count()} documents in contract and:\n")

    overdue(reg)
    outstanding(reg)

report(reg)
