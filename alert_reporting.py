import os
import base64
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib
import warnings
import smtplib
import ctypes
import tkinter as tk
from tkinter import simpledialog
from tkinter import filedialog as fd
from datetime import date, datetime, timedelta
from pretty_html_table import build_table
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# # Read in excel file
# base_path = "C:\\Users\\"
# filetypes = (('All files', '*.*'),
#              ('Excel files', '*.xlsx'))
# # show the open file dialog
# f = fd.askopenfile(initialdir=base_path, filetypes=filetypes)
# df_constants = pd.read_excel(f.name, engine="openpyxl")
# f.close()

f = __file__.rsplit("/", 2)[0] + "/alert_reporting_constants.xlsx"
if ".py" in f:
    f = __file__.rsplit("\\", 2)[0] + "\\alert_reporting_constants.xlsx"
df_constants = pd.read_excel(f, engine="openpyxl")

try:
    filepath = df_constants["var"].loc[df_constants["var_name"] == "filepath"].iloc[0]
    market_path = df_constants["var"].loc[df_constants["var_name"] == "market_path"].iloc[0]
    demand_path = df_constants["var"].loc[df_constants["var_name"] == "demand_path"].iloc[0]
    fig_path = df_constants["var"].loc[df_constants["var_name"] == "fig_path"].iloc[0]
    excel_path = df_constants["var"].loc[df_constants["var_name"] == "excel_path"].iloc[0]

    # Set constants
    week_lag = df_constants["var"].loc[df_constants["var_name"] == "week_lag"].iloc[0]
    reporting_window = df_constants["var"].loc[df_constants["var_name"] == "reporting_window"].iloc[0]
    pct_diff_threshold = df_constants["var"].loc[df_constants["var_name"] == "pct_diff_threshold"].iloc[0]
    top_cust_ind = df_constants["var"].loc[df_constants["var_name"] == "top_cust_ind"].iloc[0]

    # Set emails
    recipients = df_constants["var"].loc[df_constants["var_name"] == "to_email"].iloc[0]
    gmail_user = df_constants["var"].loc[df_constants["var_name"] == "from_email"].iloc[0]
except:
    ctypes.windll.user32.MessageBoxW(0, "You are missing a required parameter! Please check : "
                                     + f.name + " and try again.", "Missing Parameters", 0)

# Read in data tab from demand plan tool
df_data = pd.read_excel(demand_path, engine="openpyxl", sheet_name="Data")
# Manipulate data tab to remove whitespace and extraneous columns
column_header_loc = df_data[df_data.columns[0]] == df_data.columns[0]
df_data = df_data.drop(index = np.array(range(df_data.loc[column_header_loc].index[0] + 1)))
df_data = df_data[["FWeek", "WkDt", "Item", "Values", "Demand Plan", "Lag Fcst", "Parent",
                   "Invoiced Orders", "Open Orders", "AP Ship To", "PY Invoiced", "PY OOS", "OOS"]]
df_data = df_data.loc[df_data["Values"] == "Vol"]
df_data = df_data.fillna(0)
df_data["Item"] = "AP" + df_data["Item"]
df_data["WkDt"] = pd.to_datetime(df_data["WkDt"])
df_data["total_forc"] = df_data["Demand Plan"] + df_data["Lag Fcst"]
df_data["prev_year"] = df_data["PY Invoiced"] + df_data["PY OOS"]
df_data["total_orders"] = df_data["Invoiced Orders"] + df_data["Open Orders"] + df_data["OOS"]

# Get today's date and set the desired reporting range
current_date = pd.Timestamp(date.today()) - timedelta(weeks=week_lag)
reporting_range = pd.date_range(start=current_date - timedelta(weeks=reporting_window), end=current_date, freq="W-MON")
date_var = "WkDt"

########################################################################################################################
# Do the Exception Reporting
########################################################################################################################
# Filter full data set to focus on reporting range
df_comparison = df_data.loc[(df_data[date_var] >= reporting_range.min()) &
                            (df_data[date_var] <= reporting_range.max())]
# Get array of unique items and set a blank array for valid items
unique_items = pd.unique(df_comparison["Item"])
valid_items = np.array([])
for ii in range(len(unique_items)):
    # Take only items that have orders during the filtered period
    item_name = unique_items[ii]
    df_trial = df_comparison.loc[df_comparison["Item"] == item_name]
    order_forc_sum = df_trial[["total_orders", "total_forc"]].sum().sum()
    if order_forc_sum <= 0:
        continue
    else:
        valid_items = np.append(valid_items, item_name)

# Look at pct difference between forecast and actuals at item level
item_pct_diff_array = np.array([])
for ii in range(len(valid_items)):
    # Get item subset
    item_name = valid_items[ii]
    df_item = df_comparison.loc[df_comparison["Item"] == item_name]

    # Get total orders, total forecst, and difference
    total_orders = int(df_item["total_orders"].sum())
    total_forc = int(df_item["total_forc"].sum())
    total_diff = total_forc - total_orders

    # If there is a forecast, get the % difference. If not, then the difference is infinite
    if total_forc != 0:
        pct_diff = round(((total_orders - total_forc)/total_forc) * 100, 2)
    else:
        pct_diff = np.inf
    item_array = np.array([item_name, total_orders, total_forc, total_diff, pct_diff])
    item_pct_diff_array = np.append(item_pct_diff_array, item_array)

# Reform the array to convert into a dataframe
dims = [int(len(item_pct_diff_array)/len(item_array)),
        int(len(item_array))]
item_pct_diff_array = item_pct_diff_array.reshape(dims[0], dims[1])

df_item_pct_diff = pd.DataFrame(item_pct_diff_array, columns=["Item Name", "Sales", "Forecast", "Diff", "Pct Diff"])
cols = df_item_pct_diff.columns.drop("Item Name")
df_item_pct_diff[cols] = df_item_pct_diff[cols].apply(pd.to_numeric, errors='coerce')
df_item_pct_diff = df_item_pct_diff.sort_values(by="Sales", ascending=False).reset_index()
df_item_pct_diff = df_item_pct_diff.drop(columns="index")
df_item_pct_diff["Sales"] = df_item_pct_diff.apply(lambda x: "{:,}".format(x["Sales"]), axis=1)
df_item_pct_diff["Forecast"] = df_item_pct_diff.apply(lambda x: "{:,}".format(x["Forecast"]), axis=1)
df_item_pct_diff["Diff"] = df_item_pct_diff.apply(lambda x: "{:,}".format(x["Diff"]), axis=1)

#################################################
# Customers
#################################################
# Get percent difference for customer-item combinations, sorting customers by volume
unique_customers = pd.unique(df_comparison["Parent"])
sorted_volumes = df_comparison["total_orders"].groupby(df_comparison["Parent"]).sum().sort_values(ascending = False)
df_alert = pd.DataFrame()
for ii in range(len(unique_customers)):
    # Assign customer name and ignore customers where the total forecast and orders are both 0
    customer_name = unique_customers[ii]
    df_customer = df_comparison.loc[df_comparison["Parent"] == customer_name]
    if (df_customer["total_orders"].sum() == 0) and (df_customer["total_forc"].sum() == 0):
        continue
    customer_items = pd.unique(df_customer["Item"])
    sorted_items = df_customer["total_orders"].groupby(df_customer["Item"]).sum().sort_values(ascending = False)

    # Go through all the customer items
    for jj in range(len(customer_items)):
        item_name = customer_items[jj]
        df_customer_item = df_customer.loc[df_customer["Item"] == item_name]
        date_range = pd.unique(df_customer_item[date_var])

        # Create a dataframe that combines the forecast, previous year sales, and recent orders into one dataframe
        ci_sub = pd.DataFrame(date_range, columns=[date_var])
        ci_sub = pd.merge(ci_sub,
                          df_customer_item["total_forc"].groupby(df_customer_item[date_var]).sum().reset_index(),
                          on=date_var, how="left")
        ci_sub = pd.merge(ci_sub,
                          df_customer_item["prev_year"].groupby(df_customer_item[date_var]).sum().reset_index(),
                          on=date_var, how="left")
        ci_sub = pd.merge(ci_sub,
                          df_customer_item["total_orders"].groupby(df_customer_item[date_var]).sum().reset_index(),
                          on=date_var, how="left")

        # Find the percent difference between the orders and the forecast
        if ci_sub["total_orders"].sum() != 0 and ci_sub["total_forc"].sum() != 0:
            forc_order_pct_diff = (ci_sub["total_forc"].sum() - ci_sub["total_orders"].sum())/\
                                  ci_sub["total_orders"].sum()
        else:
            forc_order_pct_diff = 0

        # Find the percent difference between orders and the previous year
        if ci_sub["prev_year"].sum() != 0 and ci_sub["total_orders"].sum() != 0 and ci_sub["total_forc"].sum() != 0:
            order_prev_pct_diff = (ci_sub["total_orders"].sum() - ci_sub["prev_year"].sum()) / \
                                  ci_sub["prev_year"].sum()
        else:
            order_prev_pct_diff = 0

        if ((abs(forc_order_pct_diff) > pct_diff_threshold/100) or \
                (abs(order_prev_pct_diff) > pct_diff_threshold/100)) and \
                (customer_name in sorted_volumes.index[0:top_cust_ind]):
            total_orders = int(ci_sub["total_orders"].sum())
            total_forc = int(ci_sub["total_forc"].sum())
            total_py = int(ci_sub["prev_year"].sum())

            if abs(forc_order_pct_diff) > pct_diff_threshold/100 and \
                (abs(order_prev_pct_diff) > pct_diff_threshold/100):
                type = "both"
                forc_diff = round(forc_order_pct_diff, 2) * 100
                py_diff = round(order_prev_pct_diff, 2) * 100
            elif abs(forc_order_pct_diff) > pct_diff_threshold/100:
                type = "forecast"
                forc_diff = round(forc_order_pct_diff, 2) * 100
                py_diff = "N/A"
            else:
                type = "prev_year"
                forc_diff = "N/A"
                py_diff = round(order_prev_pct_diff, 2) * 100

            df_cust_alert_sub = pd.DataFrame([[customer_name, item_name, total_orders, total_forc, total_py,
                                              forc_diff, py_diff]],
                                             columns = ["Parent", "Item","total_orders", "total_forc", "prev_year",
                                                        "forc_diff", "py_diff"])
            df_alert = df_alert.append(df_cust_alert_sub, ignore_index=True)

df_alert["total_orders"] = df_alert.apply(lambda x: "{:,}".format(x["total_orders"]), axis=1)
df_alert["total_forc"] = df_alert.apply(lambda x: "{:,}".format(x["total_forc"]), axis=1)
df_alert["prev_year"] = df_alert.apply(lambda x: "{:,}".format(x["prev_year"]), axis=1)

########################################################################################################################
# Send the email
########################################################################################################################
# Create the necessary email environment
ROOT = tk.Tk()
ROOT.withdraw()

# Input password
# gmail_password = simpledialog.askstring(title="Password Input",
#                                   prompt="Please enter your password:", show="*")

gmail_password = base64.b64decode("VGFpZmFyITE=").decode("utf-8")

# Create message information
message = MIMEMultipart()
message['Subject'] = 'Curation Avocado Variance Report : ' + str(date.today())
message['From'] = gmail_user
message["To"] = recipients
# Create table content
date_range_text = "Date Range : " + str(reporting_window) + " Weeks - " \
                  + str(reporting_range[0].date()) + " to " + str(reporting_range[-1].date())
body_content_item = build_table(df_item_pct_diff, 'blue_light')
body_content_cust = build_table(df_alert, 'orange_light')

message.attach(MIMEText(date_range_text, "plain"))
message.attach(MIMEText(body_content_item, "html"))
message.attach(MIMEText(body_content_cust, "html"))
msg_body = message.as_string()

# Try to send the email
correct_pass = 0
try:
    while correct_pass == 0:
        try:
            smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            smtp_server.ehlo()
            smtp_server.login(gmail_user, gmail_password)
            smtp_server.sendmail(message["From"], message["To"], msg_body)
            smtp_server.close()

            correct_pass = 1
        except:
            smtp_server.close()
            gmail_password = simpledialog.askstring(title="Incorrect Password",
                                                    prompt="Please enter your password again:", show="*")
    print ("Email sent successfully!")
    ctypes.windll.user32.MessageBoxW(0, "Your Email Sent Successfully", "Email Success Confirmation", 0)
except Exception as ex:
    print ("Something went wrong….",ex)
    ctypes.windll.user32.MessageBoxW(0, "Something went wrong..." + ex, "Email Error", 0)