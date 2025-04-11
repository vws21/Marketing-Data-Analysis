import pandas as pd
import matplotlib.pyplot as plt

# read excel file
xls = pd.ExcelFile("Cleaned_Updated_Marketing_Data.xlsx")

# load the "Prospect Journey BY EMAIL"
prospect_emails_updated = xls.parse("Prospect Journey BY EMAIL")

# filter for emails with sufficient delivery numbers
valid_prospect_emails_updated = prospect_emails_updated[prospect_emails_updated["Total Delivered"] >= 30].copy()

# ensure correct data types
valid_prospect_emails_updated["Unique Open Rate"] = pd.to_numeric(valid_prospect_emails_updated["Unique Open Rate"], errors='coerce')
valid_prospect_emails_updated["Unique Click Rate"] = pd.to_numeric(valid_prospect_emails_updated["Unique Click Rate"], errors='coerce')

# identify top 5 by open and click rates
top_open_updated = valid_prospect_emails_updated.sort_values("Unique Open Rate", ascending=False).head(5)
top_click_updated = valid_prospect_emails_updated.sort_values("Unique Click Rate", ascending=False).head(5)

# combine both lists and drop duplicates
top_combined_updated = pd.concat([top_open_updated, top_click_updated]).drop_duplicates(subset=["Email Name"])

# sort for plotting
top_combined_sorted_updated = top_combined_updated.sort_values("Unique Open Rate", ascending=True)

# print(top_combined_sorted_updated)

# plot the chart
plt.figure(figsize=(12, 8))
plt.barh(top_combined_sorted_updated["Email Name"], top_combined_sorted_updated["Unique Open Rate"], label="Open Rate", alpha=0.6)
plt.barh(top_combined_sorted_updated["Email Name"], top_combined_sorted_updated["Unique Click Rate"], label="Click Rate", alpha=0.6)
plt.xlabel("Engagement Rate")
plt.title("Comparison of Open and Click Rates for Top Prospect Emails")
plt.legend()
plt.tight_layout()
plt.show()

# Load the "Admit Journey BY EMAIL"
admit_emails = xls.parse("Admit Journey BY EMAIL")

# remove rows with missing email names
admit_emails_cleaned = admit_emails.dropna(subset=["Email Name"])

# filter for emails with sufficient delivery numbers
valid_admit_emails_cleaned = admit_emails_cleaned[admit_emails_cleaned["Total Delivered"] >= 30].copy()

# top open and click selections
top_open_admit_cleaned = valid_admit_emails_cleaned.sort_values("Unique Open Rate", ascending=False).head(5)
top_click_admit_cleaned = valid_admit_emails_cleaned.sort_values("Unique Click Rate", ascending=False).head(5)

# combine and deduplicate
top_combined_admit_cleaned = pd.concat([top_open_admit_cleaned, top_click_admit_cleaned]).drop_duplicates(subset=["Email Name"])
top_combined_admit_sorted_cleaned = top_combined_admit_cleaned.sort_values("Unique Open Rate", ascending=True)

# print(top_combined_admit_sorted_cleaned)

# plot the chart
plt.figure(figsize=(12, 8))
plt.barh(top_combined_admit_sorted_cleaned["Email Name"], top_combined_admit_sorted_cleaned["Unique Open Rate"], label="Open Rate", alpha=0.6)
plt.barh(top_combined_admit_sorted_cleaned["Email Name"], top_combined_admit_sorted_cleaned["Unique Click Rate"], label="Click Rate", alpha=0.6)
plt.xlabel("Engagement Rate")
plt.title("Comparison of Open and Click Rates for Top Admit Emails")
plt.legend()
plt.tight_layout()
plt.show()