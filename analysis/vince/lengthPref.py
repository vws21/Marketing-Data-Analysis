import pandas as pd
import matplotlib.pyplot as plt

# read excel file
xls = pd.ExcelFile('Cleaned_Updated_Marketing_Data.xlsx')

# load email performance data
prospect_perf = pd.read_excel(xls, sheet_name='Prospect Journey BY EMAIL')
admit_perf = pd.read_excel(xls, sheet_name='Admit Journey BY EMAIL')
prospect_perf['Stage'] = 'Prospective'
admit_perf['Stage'] = 'Admitted'

# extract columns
email_metrics = pd.concat([
    prospect_perf[['Email Name', 'Unique Open Rate', 'Unique Click Rate', 'Stage']],
    admit_perf[['Email Name', 'Unique Open Rate', 'Unique Click Rate', 'Stage']]
], ignore_index=True)

# load email content data
prospect_emails = pd.read_excel(xls, sheet_name='Prospect Emails (Current)')
admit_emails = pd.read_excel(xls, sheet_name='Admit Emails (Current)')
prospect_emails['Stage'] = 'Prospective'
admit_emails['Stage'] = 'Admitted'

# compute word count
prospect_emails['Word Count'] = prospect_emails['Email Body Text'].str.split().apply(len)
admit_emails['Word Count'] = admit_emails['Email Body Text'].str.split().apply(len)

# extract needed fields
email_text_data = pd.concat([
    prospect_emails[['Email Name', 'Word Count', 'Stage']],
    admit_emails[['Email Name', 'Word Count', 'Stage']]
], ignore_index=True)

# merge word count with engagement metrics
merged = pd.merge(email_metrics, email_text_data, on=['Email Name', 'Stage'], how='inner')

# bin word counts
bins = [0, 100, 200, 300, 400, 500, 1000]
labels = ['0-100', '101-200', '201-300', '301-400', '401-500', '501+']
merged['Word Count Bin'] = pd.cut(merged['Word Count'], bins=bins, labels=labels, right=False)

# average engagement rates by bin
engagement = merged.groupby(['Stage', 'Word Count Bin'])[
    ['Unique Open Rate', 'Unique Click Rate']
].mean().reset_index()

# plot
fig, ax = plt.subplots(figsize=(10, 6))
for stage in engagement['Stage'].unique():
    data = engagement[engagement['Stage'] == stage]
    ax.plot(data['Word Count Bin'], data['Unique Open Rate'], marker='o', label=f'{stage} - Open Rate')
    ax.plot(data['Word Count Bin'], data['Unique Click Rate'], marker='x', linestyle='--', label=f'{stage} - Click Rate')

ax.set_title('Email Engagement by Word Count Bin and Student Stage')
ax.set_xlabel('Email Word Count Bin')
ax.set_ylabel('Engagement Rate')
ax.legend()
ax.grid(True)
plt.tight_layout()
plt.show()