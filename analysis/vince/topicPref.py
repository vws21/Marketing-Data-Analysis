import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# read excel file
xls = pd.ExcelFile("Cleaned_Updated_Marketing_Data.xlsx")

# load relevant sheets
prospect_emails = pd.read_excel(xls, sheet_name='Prospect Journey BY EMAIL')
admit_emails = pd.read_excel(xls, sheet_name='Admit Journey BY EMAIL')
subject_keys = pd.read_excel(xls, sheet_name='Journey Subject Keys')

# merge subject lines with email engagement data
prospect = pd.merge(prospect_emails, subject_keys, on='Key', how='left').dropna(subset=['Subject Line'])
admit = pd.merge(admit_emails, subject_keys, on='Key', how='left').dropna(subset=['Subject Line'])

# topic classification function
def classify_topic(subject):
    subject = subject.lower()
    if "career" in subject or "job" in subject or "alumni" in subject:
        return "Career"
    elif "application" in subject or "apply" in subject:
        return "Application"
    elif "faculty" in subject or "professor" in subject or "research" in subject:
        return "Faculty/Research"
    elif "pittsburgh" in subject or "city" in subject or "campus" in subject or "life" in subject:
        return "Campus Life"
    elif "event" in subject or "session" in subject or "webinar" in subject:
        return "Events"
    elif "advising" in subject or "help" in subject or "support" in subject:
        return "Support/Advising"
    else:
        return "General Info"

# apply topic classification
prospect['Topic'] = prospect['Subject Line'].apply(classify_topic)
admit['Topic'] = admit['Subject Line'].apply(classify_topic)

# group and calculate mean engagement metrics
prospect_summary = prospect.groupby('Topic')[['Unique Open Rate', 'Unique Click Rate']].mean().reset_index()
admit_summary = admit.groupby('Topic')[['Unique Open Rate', 'Unique Click Rate']].mean().reset_index()
prospect_summary['Stage'] = 'Prospect'
admit_summary['Stage'] = 'Admit'

# combine and sort
topic_engagement = pd.concat([prospect_summary, admit_summary], ignore_index=True)
print(topic_engagement)

# melt for combined bar charts
melted = topic_engagement.melt(id_vars=['Topic', 'Stage'],
                               value_vars=['Unique Open Rate', 'Unique Click Rate'],
                               var_name='Metric', value_name='Rate')

# plot: stacked vertically (row layout), x labels fixed
sns.set(style="whitegrid")
g = sns.catplot(
    data=melted,
    kind="bar",
    x="Topic", y="Rate", hue="Stage",
    row="Metric", height=5, aspect=2,
    palette="Set2", sharey=False
)

# fix titles and labels
g.set_titles("{row_name}")
g.set_axis_labels("Email Topic", "Engagement Rate")

# rotate and re-enable x-axis labels for all rows
for ax in g.axes.flat:
    ax.tick_params(labelbottom=True)
    ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')

plt.tight_layout()
plt.show()