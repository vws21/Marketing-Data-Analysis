import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# read excel file
df = pd.read_excel('Cleaned_Updated_Marketing_Data.xlsx', sheet_name='Single Send Summary')

# print the dataframe
print(df)

df["Date and Time Sent"] = pd.to_datetime(df["Date and Time Sent"])
df["Hour Sent"] = df["Date and Time Sent"].dt.hour
df["Day of Week"] = df["Date and Time Sent"].dt.day_name()
df["Month"] = df["Date and Time Sent"].dt.month_name()

print(df["Hour Sent"])

# set style
sns.set_style("whitegrid")

# open rate by hour
plt.figure(figsize=(7, 5))
sns.boxplot(x=df["Hour Sent"], y=df["Open Rate"])
plt.xlabel("Hour of the Day")
plt.ylabel("Open Rate")
plt.title("Email Open Rate by Hour of the Day")
plt.xticks(range(0, 19, 1))
plt.show()

# open rate by day of week
plt.figure(figsize=(10, 5))
sns.boxplot(x=df["Day of Week"], y=df["Open Rate"], order=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
plt.xlabel("Day of the Week")
plt.ylabel("Open Rate")
plt.title("Email Open Rate by Day of the Week")
plt.show()

# ctr by month
plt.figure(figsize=(10, 5))
sns.boxplot(x=df["Month"], y=df["Unique Click-Through Rate"], order=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])
plt.xlabel("Month")
plt.ylabel("Click-Through Rate")
plt.title("Email Click-Through Rate by Month")
plt.xticks(rotation=45)
plt.show()

# heatmap: open rate by hour and day of the week
heatmap_data = df.pivot_table(values="Open Rate", index="Day of Week", columns="Hour Sent", aggfunc="mean")
heatmap_data = heatmap_data.reindex(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
plt.figure(figsize=(12, 6))
sns.heatmap(heatmap_data, cmap="coolwarm", annot=True, fmt=".2f")
plt.xlabel("Hour of the Day")
plt.ylabel("Day of the Week")
plt.title("Heatmap of Open Rates by Hour and Day of the Week")
plt.show()