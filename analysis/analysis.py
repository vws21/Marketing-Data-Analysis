import pandas as pd
import matplotlib.pyplot as plt
from sklearn.preprocessing import LabelEncoder

# load the new Excel workbook
file_path = "Cleaned_Updated_Marketing_Data.xlsx"
xls = pd.ExcelFile(file_path)

# ANALYSIS FOR: What is the engagement of longer vs shorter subject lines?
# Define relevant sheets
sheets_to_analyze = [
    'Prospect Journey BY EMAIL',
    'Admit Journey BY EMAIL',
    'Info Session BY EMAIL',
    'Single Send BY EMAIL'
]

# Load data from these sheets
dataframes = {sheet: xls.parse(sheet) for sheet in sheets_to_analyze}

# Load subject keys for subject lines
subject_keys_df = xls.parse("Journey Subject Keys")

# Clean and standardize "Key" column for merging
subject_keys_df["Key"] = subject_keys_df["Key"].astype(str).str.strip().astype(int)
subject_keys_df["Subject Length"] = subject_keys_df["Subject Line"].astype(str).apply(len)

# Merge subject line data with each email dataset
for sheet in sheets_to_analyze:
    if "Key" in dataframes[sheet].columns:
        dataframes[sheet]["Key"] = dataframes[sheet]["Key"].astype(str).str.strip().astype(int)
        dataframes[sheet] = dataframes[sheet].merge(subject_keys_df[["Key", "Subject Length"]], on="Key", how="left")

# Extract relevant metrics for analysis
engagement_data = []

for sheet, df in dataframes.items():
    if "Unique Open Rate" in df.columns and "Unique Click Rate" in df.columns and "Subject Length" in df.columns:
        df_filtered = df[["Subject Length", "Unique Open Rate", "Unique Click Rate"]].dropna()
        df_filtered["Source"] = sheet  # Track which dataset the data came from
        engagement_data.append(df_filtered)

# Combine all sources into one DataFrame
engagement_df = pd.concat(engagement_data, ignore_index=True)
engagement_df["Unique Open Rate"] = pd.to_numeric(engagement_df["Unique Open Rate"], errors="coerce")
engagement_df["Unique Click Rate"] = pd.to_numeric(engagement_df["Unique Click Rate"], errors="coerce")

open_rate_plot_path = "/Users/nikhitavysyaraju/Downloads/INFSC 1740/subject_length_vs_open_rate.png"
click_rate_plot_path = "/Users/nikhitavysyaraju/Downloads/INFSC 1740/subject_length_vs_click_rate.png"

# Scatter plot for subject length vs unique open rate
plt.figure(figsize=(8, 6))
plt.scatter(engagement_df["Subject Length"], engagement_df["Unique Open Rate"], alpha=0.5)
plt.xlabel("Subject Line Length")
plt.ylabel("Unique Open Rate")
plt.title("Subject Line Length vs Unique Open Rate")
plt.grid(True)
#plt.show()
plt.savefig(open_rate_plot_path)  # Save figure
plt.close()

# Scatter plot for subject length vs unique click rate
plt.figure(figsize=(8, 6))
plt.scatter(engagement_df["Subject Length"], engagement_df["Unique Click Rate"], alpha=0.5, color='r')
plt.xlabel("Subject Line Length")
plt.ylabel("Unique Click Rate")
plt.title("Subject Line Length vs Unique Click Rate")
plt.grid(True)
#plt.show()
plt.savefig(click_rate_plot_path)  # Save figure
plt.close()

# Drop rows with NaNs just to be safe
engagement_df_clean = engagement_df.dropna(subset=["Subject Length", "Unique Open Rate", "Unique Click Rate"])

# Calculate correlation using pandas
open_corr = engagement_df_clean["Subject Length"].corr(engagement_df_clean["Unique Open Rate"])
click_corr = engagement_df_clean["Subject Length"].corr(engagement_df_clean["Unique Click Rate"])

print(f"Correlation between Subject Length and Unique Open Rate: {open_corr:.4f}")
print(f"Correlation between Subject Length and Unique Click Rate: {click_corr:.4f}")

# ANALYSIS FOR: What is the engagement of longer vs. shorter emails? 
# Sheets that include email body text
email_body_sheets = [
    'Admit Emails (Current)',
    'Info Session (Current)',
    'Single Send (Content)',
    'Prospect Emails (Current)'
]

# Load those sheets
body_dataframes = {sheet: xls.parse(sheet) for sheet in email_body_sheets}

# Extract Key + Email Body Length from email_body_sheets
email_body_info = []

for sheet in email_body_sheets:
    df = body_dataframes[sheet]
    if "Key" in df.columns and "Email Body Text" in df.columns:
        df = df.dropna(subset=["Key", "Email Body Text"])
        df["Key"] = df["Key"].astype(str).str.strip().astype(int)
        df["Email Length"] = df["Email Body Text"].astype(str).apply(len)
        df["Body Source"] = sheet
        email_body_info.append(df[["Key", "Email Length", "Body Source"]])

email_body_df = pd.concat(email_body_info, ignore_index=True)

# Extract engagement info from the already-loaded `dataframes` from sheets_to_analyze
engagement_info = []

for sheet in sheets_to_analyze:
    df = dataframes[sheet]
    if "Key" in df.columns and "Unique Open Rate" in df.columns and "Unique Click Rate" in df.columns:
        df = df.dropna(subset=["Key", "Unique Open Rate", "Unique Click Rate"])
        df["Key"] = df["Key"].astype(str).str.strip().astype(int)
        df["Engagement Source"] = sheet
        df["Unique Open Rate"] = pd.to_numeric(df["Unique Open Rate"], errors="coerce")
        df["Unique Click Rate"] = pd.to_numeric(df["Unique Click Rate"], errors="coerce")
        engagement_info.append(df[["Key", "Unique Open Rate", "Unique Click Rate", "Engagement Source"]])

engagement_df = pd.concat(engagement_info, ignore_index=True)

# Merge on Key
email_engagement_df = pd.merge(email_body_df, engagement_df, on="Key", how="inner")

# Drop rows with NaNs just to be safe
email_engagement_df_clean = email_engagement_df.dropna(subset=["Email Length", "Unique Open Rate", "Unique Click Rate"])

# Calculate correlation using pandas
open_corr_body = email_engagement_df_clean["Email Length"].corr(email_engagement_df_clean["Unique Open Rate"])
click_corr_body = email_engagement_df_clean["Email Length"].corr(email_engagement_df_clean["Unique Click Rate"])

print("\n=== Correlation Between Email Body Length and Engagement ===")
print(f"Correlation between Email Length and Unique Open Rate: {open_corr_body:.4f}")
print(f"Correlation between Email Length and Unique Click Rate: {click_corr_body:.4f}")

# Visualization

plt.figure(figsize=(8, 6))
plt.scatter(email_engagement_df_clean["Email Length"], email_engagement_df_clean["Unique Open Rate"], alpha=0.5)
plt.xlabel("Email Body Length (Characters)")
plt.ylabel("Unique Open Rate")
plt.title("Email Length vs Unique Open Rate")
plt.grid(True)
plt.savefig("analysis/email_length_vs_open_rate.png")
plt.close()

plt.figure(figsize=(8, 6))
plt.scatter(email_engagement_df_clean["Email Length"], email_engagement_df_clean["Unique Click Rate"], alpha=0.5, color='green')
plt.xlabel("Email Body Length (Characters)")
plt.ylabel("Unique Click Rate")
plt.title("Email Length vs Unique Click Rate")
plt.grid(True)
plt.savefig("analysis/email_length_vs_click_rate.png")
plt.close()


# ANALYSIS FOR: What is the engagement in emails by topic?  
# IE: application vs. career development resources vs. the city of Pittsburgh, etc.?
# === Add topic labels based on keywords ===
def classify_topic(text):
    text = text.lower()
    if "application" in text or "apply" in text:
        return "Application"
    elif "career" in text or "job" in text or "resume" in text:
        return "Career Development"
    elif "pittsburgh" in text or "city" in text:
        return "City of Pittsburgh"
    elif "info session" in text or "webinar" in text:
        return "Information Session"
    else:
        return "Other"

# Load full body text again for labeling
full_body_info = []

for sheet in email_body_sheets:
    df = body_dataframes[sheet]
    if "Key" in df.columns and "Email Body Text" in df.columns:
        df = df.dropna(subset=["Key", "Email Body Text"])
        df["Key"] = df["Key"].astype(str).str.strip().astype(int)
        df["Email Body Text"] = df["Email Body Text"].astype(str)
        df["Topic"] = df["Email Body Text"].apply(classify_topic)
        df = df[["Key", "Topic"]]
        full_body_info.append(df)

topics_df = pd.concat(full_body_info, ignore_index=True)

# Merge topic into your engagement DataFrame
email_engagement_df_topic = pd.merge(email_engagement_df_clean, topics_df, on="Key", how="left")

label_encoder = LabelEncoder()
email_engagement_df_topic["Topic_Code"] = label_encoder.fit_transform(email_engagement_df_topic["Topic"])

# Calculate correlation using pandas
open_corr_topic = email_engagement_df_topic["Topic_Code"].corr(email_engagement_df_topic["Unique Open Rate"])
click_corr_topic = email_engagement_df_topic["Topic_Code"].corr(email_engagement_df_topic["Unique Click Rate"])

print("\n=== Correlation Between Email Topic (Encoded) and Engagement ===")
print(f"Correlation between Topic (encoded) and Unique Open Rate: {open_corr_topic:.4f}")
print(f"Correlation between Topic (encoded) and Unique Click Rate: {click_corr_topic:.4f}")


# Group by Topic and calculate mean engagement
topic_engagement_summary = email_engagement_df_topic.groupby("Topic")[["Unique Open Rate", "Unique Click Rate"]].mean().reset_index()

print("\n=== Average Engagement by Topic ===")
print(topic_engagement_summary)

# Bar chart for Open Rate by Topic
plt.figure(figsize=(10, 6))
plt.bar(topic_engagement_summary["Topic"], topic_engagement_summary["Unique Open Rate"])
plt.xlabel("Topic")
plt.ylabel("Average Unique Open Rate")
plt.title("Average Open Rate by Email Topic")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("analysis/open_rate_by_topic.png")
plt.close()

# Bar chart for Click Rate by Topic
plt.figure(figsize=(10, 6))
plt.bar(topic_engagement_summary["Topic"], topic_engagement_summary["Unique Click Rate"], color='orange')
plt.xlabel("Topic")
plt.ylabel("Average Unique Click Rate")
plt.title("Average Click Rate by Email Topic")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("analysis/click_rate_by_topic.png")
plt.close()


# ANALYSIS FOR: What is the engagement of emails that have no personalization fields vs. 1, 2, or more personalization fields? 
# (Exclude Information Session emails in this analysis)
# Define the sheets that include full email bodies
body_sheets = [
    'Admit Emails (Current)',
    'Single Send (Content)',
    'Prospect Emails (Current)'
]

# Load those sheets
body_dataframes = {sheet: xls.parse(sheet) for sheet in body_sheets}

# Define the personalization field keywords again
personalization_keywords = ["first name", "last name", "program", "degree", "email", "location", "school"]

# Helper function to count personalization fields in an email
def count_personalization_fields(text):
    text = str(text).lower()
    return sum(1 for keyword in personalization_keywords if keyword in text)

# Process data and collect for analysis, excluding "Info Session (Current)"
personalization_records = []

for sheet in body_sheets:
    if sheet == "Info Session (Current)":
        continue  # Skip Info Session
    df = body_dataframes[sheet]
    if "Key" in df.columns and "Email Body Text" in df.columns:
        df = df.dropna(subset=["Key", "Email Body Text"])
        df["Key"] = df["Key"].astype(str).str.strip().astype(int)
        df["Personalization Count"] = df["Email Body Text"].apply(count_personalization_fields)
        df["Body Source"] = sheet
        personalization_records.append(df[["Key", "Personalization Count", "Body Source"]])

# Combine all personalization info
personalization_df = pd.concat(personalization_records, ignore_index=True)

# Load engagement info again from previously used sheets
engagement_sheets = ['Prospect Journey BY EMAIL', 'Admit Journey BY EMAIL', 'Single Send BY EMAIL']
engagement_info = []

for sheet in engagement_sheets:
    df = dataframes[sheet]
    if "Key" in df.columns and "Unique Open Rate" in df.columns and "Unique Click Rate" in df.columns:
        df = df.dropna(subset=["Key", "Unique Open Rate", "Unique Click Rate"])
        df["Key"] = df["Key"].astype(str).str.strip().astype(int)
        df["Engagement Source"] = sheet
        df["Unique Open Rate"] = pd.to_numeric(df["Unique Open Rate"], errors="coerce")
        df["Unique Click Rate"] = pd.to_numeric(df["Unique Click Rate"], errors="coerce")
        engagement_info.append(df[["Key", "Unique Open Rate", "Unique Click Rate", "Engagement Source"]])

# Combine engagement data
engagement_df = pd.concat(engagement_info, ignore_index=True)

# Merge with personalization data on Key
merged_df = pd.merge(personalization_df, engagement_df, on="Key", how="inner")

# Group personalization counts into bins (0, 1, 2, 3+)
def categorize_fields(n):
    if n == 0:
        return "0"
    elif n == 1:
        return "1"
    elif n == 2:
        return "2"
    else:
        return "3+"

merged_df["Personalization Group"] = merged_df["Personalization Count"].apply(categorize_fields)

# Aggregate average engagement by personalization group
personalization_summary = merged_df.groupby("Personalization Group")[["Unique Open Rate", "Unique Click Rate"]].mean().reset_index()

print(personalization_summary)

# Drop any rows with NaNs in engagement or personalization for correlation analysis
merged_clean_df = merged_df.dropna(subset=["Personalization Count", "Unique Open Rate", "Unique Click Rate"])

# Calculate Pearson correlations
open_corr = merged_clean_df["Personalization Count"].corr(merged_clean_df["Unique Open Rate"])
click_corr = merged_clean_df["Personalization Count"].corr(merged_clean_df["Unique Click Rate"])

open_corr, click_corr

print("\n=== Correlation Between Personalizaiton and Engagement ===")
print(f"Correlation between Number of Personalization Fields and Unique Open Rate: {open_corr:.4f}")
print(f"Correlation between Number of Personalization Fields and Unique Click Rate: {click_corr:.4f}")

# Bar chart: Average Unique Open Rate by Personalization Group
plt.figure(figsize=(8, 6))
plt.bar(personalization_summary["Personalization Group"], personalization_summary["Unique Open Rate"])
plt.xlabel("Number of Personalization Fields")
plt.ylabel("Average Unique Open Rate")
plt.title("Open Rate by Number of Personalization Fields")
plt.grid(True, axis='y')
plt.tight_layout()
plt.savefig("analysis/open_rate_by_personalization.png")
plt.close()

# Bar chart: Average Unique Click Rate by Personalization Group
plt.figure(figsize=(8, 6))
plt.bar(personalization_summary["Personalization Group"], personalization_summary["Unique Click Rate"], color='orange')
plt.xlabel("Number of Personalization Fields")
plt.ylabel("Average Unique Click Rate")
plt.title("Click Rate by Number of Personalization Fields")
plt.grid(True, axis='y')
plt.tight_layout()
plt.savefig("analysis/click_rate_by_personalization.png")
plt.close()

# Return the paths to the saved plots
["analysis/open_rate_by_personalization.png", "analysis/click_rate_by_personalization.png"]