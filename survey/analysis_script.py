import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load data
df = pd.read_excel('survey.xlsx')

# Remove rows that contain question text (they have the actual question in the response)
df = df[~df['Q3'].str.contains('how often', case=False, na=False)]
df = df[~df['Q22'].str.contains('how often', case=False, na=False)]
df = df[~df['Q23'].str.contains('how often', case=False, na=False)]
df = df[~df['Q17'].str.contains(
    'what is your preference', case=False, na=False)]

# Add this color mapping at the top of the file
LIKERT_COLORS = {
    'Strongly Agree': '#2ecc71',
    'Strongly agree': '#2ecc71',
    'Somewhat Agree': '#a0d8b3',
    'Agree': '#a0d8b3',
    'Neutral': '#95a5a6',
    'Neither agree nor disagree': '#95a5a6',
    'Somewhat Disagree': '#f26f61',
    'Disagree': '#f26f61',
    'Strongly Disagree': '#e74c3c'
}

# Convert Likert scale columns to numeric, handling text responses


def convert_likert(x):
    if pd.isna(x):
        return np.nan
    # Handle text responses
    likert_map = {
        'Strongly disagree': 1,
        'Disagree': 2,
        'Neither agree nor disagree': 3,
        'Agree': 4,
        'Strongly agree': 5
    }
    return likert_map.get(str(x).strip(), x)


# Convert Likert scale columns
likert_items = ['Q13_1', 'Q24_3', 'Q24_1', 'Q27']
for col in likert_items:
    df[col] = df[col].apply(convert_likert)
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 1. Check Data Types and Basic Distribution
print("=== Data Types ===")
print(df.dtypes)

# Unique values for categorical questions
cat_questions = ['Q3', 'Q22', 'Q23', 'Q17']
for q in cat_questions:
    print(f"\nUnique values for {q}:")
    print(df[q].value_counts(dropna=False))

# Modify the analyze_likert function to include visualization


def analyze_likert(series, question_text):
    valid_data = series.dropna()
    if len(valid_data) == 0:
        return pd.DataFrame({
            'Mean': [np.nan],
            'Std Dev': [np.nan],
            '95% CI Lower': [np.nan],
            '95% CI Upper': [np.nan],
            'Count': [0]
        })

    # Create visualization
    response_map = {
        1: 'Strongly Disagree',
        2: 'Disagree',
        3: 'Neutral',
        4: 'Agree',
        5: 'Strongly Agree'
    }
    response_counts = valid_data.map(response_map).value_counts()

    plt.figure(figsize=(10, 6))
    response_counts.sort_index(ascending=False).plot(
        kind='barh',
        color=[LIKERT_COLORS[x] for x in response_counts.index]
    )
    plt.title(f'Responses to: {question_text}')
    plt.xlabel('Number of Responses')
    plt.ylabel('Response')
    plt.tight_layout()
    plt.savefig(f'{series.name}_distribution.png')
    plt.close()

    # Calculate statistics
    stats = {
        'Mean': [valid_data.mean()],
        'Std Dev': [valid_data.std()],
        '95% CI Lower': [valid_data.mean() - 1.96*valid_data.std()/np.sqrt(len(valid_data))],
        '95% CI Upper': [valid_data.mean() + 1.96*valid_data.std()/np.sqrt(len(valid_data))],
        'Count': [len(valid_data)]
    }
    return pd.DataFrame(stats)


# Update the Likert analysis section
likert_questions = {
    'Q13_1': "I found the HelpMe system was an effective single location to go to for the help I needed.",
    'Q24_3': "A single place for help (office hours, chatbot, Anytime Questions) is beneficial.",
    'Q24_1': "I prefer email over asking Anytime Questions via HelpMe.",
    'Q27': "I asked more questions to the chatbot or Anytime Questions than I sent emails in the course."
}

likert_results = {}
for item, question_text in likert_questions.items():
    print(f"\nAnalyzing {item}: {question_text}")
    likert_results[item] = analyze_likert(df[item], question_text)
    print(likert_results[item])
    print("\nResponse distribution:")
    print(df[item].value_counts().sort_index())

# Multiple-select analysis (Q15)
# Assuming Q15 responses are comma-separated strings
if not df['Q15'].isna().all():
    q15_options = df['Q15'].dropna().str.split(
        ',', expand=True).stack().str.strip().value_counts()
    q15_summary = pd.DataFrame({'Count': q15_options})
    print("\nQ15 Barriers to Office Hours:")
    print(q15_summary)

# 2. Single Location & Reduced Barriers
single_location = pd.DataFrame({
    'Q13_1': df['Q13_1'].value_counts().sort_index(),
    'Q24_3': df['Q24_3'].value_counts().sort_index()
})
print("\nSingle Location Responses:")
print(single_location)

# Q17 preference analysis
print("\nCurrent Help Preferences (Q17):")
print(df['Q17'].value_counts())

# 3. Frequent Usage Analysis
usage_columns = ['Q3', 'Q22', 'Q23']
usage_summary = pd.DataFrame(
    {col: df[col].value_counts() for col in usage_columns})
print("\nUsage Patterns:")
print(usage_summary)

# 4. Compare Traditional Methods
# Previous vs current preference
preference_shift = pd.crosstab(df['Q16'], df['Q17'])
print("\nPreference Shift (Q16 vs Q17):")
print(preference_shift)

# 5. Optional Cross-Analyses
high_satisfaction = df[df['Q13_1'] >= 4]
high_sat_usage = high_satisfaction[['Q22', 'Q23']].apply(
    pd.Series.value_counts)
print("\nHigh Satisfaction User Patterns:")
print(high_sat_usage)

# Output results to CSVs
likert_results['Q13_1'].to_csv('q13_1_summary.csv')
if 'q15_summary' in locals():
    q15_summary.to_csv('q15_summary.csv')
usage_summary.to_csv('usage_summary.csv')
preference_shift.to_csv('preference_shift.csv')

# Print sample interpretations
print("\nInterpretation Examples:")
print(f"- Q13_1 Mean: {likert_results['Q13_1']
      ['Mean'].iloc[0]:.2f} suggests overall agreement")
print(f"- {q15_summary.idxmax()[0]} was the most common office hour barrier")
print(f"- {preference_shift.sum().idxmax()
           } is the most common current help preference")
