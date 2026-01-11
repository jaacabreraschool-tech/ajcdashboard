import pandas as pd

# -----------------------------
# Load dataset
# -----------------------------
df = pd.read_excel("HR Cleaned Data 01.09.26.xlsx")

# Extract reporting year
df['Year'] = pd.to_datetime(df['Calendar Year']).dt.year

# -----------------------------
# Active employees
# -----------------------------
active_df = df[df['Resignee Checking '].isin(['ACTIVE', 'Active'])]
active_df = active_df[active_df['Year'].between(2020, 2025)]

# Headcount per year
headcount_per_year = (
    active_df.groupby('Year')['Full Name']
    .nunique()
    .reset_index(name='Headcount')
)

# Age distribution (✅ includes Generation now)
age_distribution = (
    active_df.groupby(['Year','Age','Generation'])['Full Name']
    .nunique()
    .reset_index(name='Count')
)

# Generation distribution
generation_distribution = (
    active_df.groupby(['Year','Generation'])['Full Name']
    .nunique()
    .reset_index(name='Count')
)

# Gender diversity
gender_distribution = (
    active_df.groupby(['Year','Gender','Position/Level'])['Full Name']
    .nunique()
    .reset_index(name='Count')
)

# Tenure analysis
active_df['YearJoined'] = pd.to_datetime(active_df['Year Joined']).dt.year
active_df['Tenure'] = active_df['Year'] - active_df['YearJoined']

tenure_distribution = (
    active_df.groupby(['Year','YearJoined','Tenure'])['Full Name']
    .nunique()
    .reset_index(name='Count')
)

# -----------------------------
# Resignation trends
# -----------------------------
leaver_df = df[df['Resignee Checking '].isin(['Leaver','LEAVER'])]
leaver_df = leaver_df[leaver_df['Year'].between(2020, 2025)]
leaver_df['YearJoined'] = pd.to_datetime(leaver_df['Year Joined']).dt.year
leaver_df['Tenure'] = leaver_df['Year'] - leaver_df['YearJoined']

resignation_trends = (
    leaver_df.groupby(['Year','YearJoined','Tenure'])['Full Name']
    .nunique()
    .reset_index(name='LeaverCount')
)

resignation_trends = resignation_trends.merge(headcount_per_year, on='Year', how='left')
resignation_trends['AttritionRate'] = (resignation_trends['LeaverCount'] / resignation_trends['Headcount']) * 100

# -----------------------------
# Retention by Cohort
# -----------------------------
df['YearJoined'] = pd.to_datetime(df['Year Joined']).dt.year
cohort_size = (
    df.groupby('YearJoined')['Full Name']
    .nunique()
    .reset_index(name='CohortSize')
)

retention_cohort = active_df[
    ['YearJoined', 'Generation', 'Position/Level', 'Full Name']
].copy()

retention_cohort = retention_cohort.merge(cohort_size, on='YearJoined', how='left')

retention_summary = (
    retention_cohort.groupby(['YearJoined','Generation','Position/Level'])
    .agg(RetainedCount=('Full Name','nunique'), CohortSize=('CohortSize','first'))
    .reset_index()
)
retention_summary['RetentionRate'] = (retention_summary['RetainedCount'] / retention_summary['CohortSize']) * 100

# -----------------------------
# Promotion & Transfer tracking
# -----------------------------
promo_transfer_df = df[df['Promotion & Transfer'] == 1]
promo_transfer_df = promo_transfer_df[promo_transfer_df['Year'].between(2020, 2025)]

promo_transfer_df['YearJoined'] = pd.to_datetime(promo_transfer_df['Year Joined']).dt.year
promo_transfer_df['Tenure'] = promo_transfer_df['Year'] - promo_transfer_df['YearJoined']

promotion_transfer_tracking = (
    promo_transfer_df.groupby(['Year','Position/Level','Tenure'])['Full Name']
    .nunique()
    .reset_index(name='Count')
)

promotion_transfer_tracking = promotion_transfer_tracking.merge(headcount_per_year, on='Year', how='left')
promotion_transfer_tracking['Rate'] = (promotion_transfer_tracking['Count'] / promotion_transfer_tracking['Headcount']) * 100

# -----------------------------
# Duplicate name check
# -----------------------------
duplicates = df[df.duplicated(subset=['YearJoined','Full Name'], keep=False)]
duplicates = duplicates.sort_values(['YearJoined','Full Name'])

# -----------------------------
# Satisfaction heatmap
# -----------------------------
likert_columns = [
    'Corporate Culture', 'Job Satisfaction', 'Pay/Benefits', 'Job Content and Design',
    'Management', 'Respect', 'Innovation', 'Career', 'Work/Life',
    'Leadership', 'Communication', 'Appraisals'
]

likert_per_year_percentage = df.groupby('Year')[likert_columns].mean().reset_index()
likert_per_year_count = df.groupby('Year')[likert_columns].count().reset_index()

# -----------------------------
# Engagement Index
# -----------------------------
engagement_columns = ['Respect', 'Career', 'Communication']
df['Engagement Index'] = df[engagement_columns].mean(axis=1)

engagement_index_per_year = (
    df.groupby('Year')['Engagement Index']
    .mean()
    .reset_index(name='Engagement Score')
)

# -----------------------------
# Driver Analysis
# -----------------------------
df['Resignee_Binary'] = df['Resignee Checking '].apply(lambda x: 1 if str(x).upper() in ['LEAVER'] else 0)

driver_resignation_yearly = []
driver_promotion_yearly = []

for year, group in df.groupby('Year'):
    corr_resignation = group[likert_columns + ['Resignee_Binary']].corr()['Resignee_Binary'].drop('Resignee_Binary')
    corr_resignation = corr_resignation.reset_index()
    corr_resignation.columns = ['Category', 'Correlation with Resignation']
    corr_resignation['Year'] = year
    driver_resignation_yearly.append(corr_resignation)

    corr_promotion = group[likert_columns + ['Promotion & Transfer']].corr()['Promotion & Transfer'].drop('Promotion & Transfer')
    corr_promotion = corr_promotion.reset_index()
    corr_promotion.columns = ['Category', 'Correlation with Promotion']
    corr_promotion['Year'] = year
    driver_promotion_yearly.append(corr_promotion)

driver_resignation_df = pd.concat(driver_resignation_yearly, ignore_index=True)
driver_promotion_df = pd.concat(driver_promotion_yearly, ignore_index=True)

promotion_predictors_df = driver_promotion_df.copy()
promotion_predictors_df.rename(columns={'Correlation with Promotion': 'Promotion Predictor Strength'}, inplace=True)

# -----------------------------
# Engagement vs Retention
# -----------------------------
engagement_retention = (
    df.groupby(['Year','Resignee Checking '])['Engagement Index']
    .mean()
    .reset_index(name='Avg Engagement Score')
)

satisfaction_retention = (
    df.groupby(['Year','Resignee Checking '])[likert_columns]
    .mean()
    .reset_index()
)

# -----------------------------
# Merge headcount
# -----------------------------
age_distribution = age_distribution.merge(headcount_per_year, on='Year')
generation_distribution = generation_distribution.merge(headcount_per_year, on='Year')
gender_distribution = gender_distribution.merge(headcount_per_year, on='Year')
tenure_distribution = tenure_distribution.merge(headcount_per_year, on='Year')

# -----------------------------
# Save all outputs to Excel
# -----------------------------
with pd.ExcelWriter('HR_Analysis_Output.xlsx') as writer:
    age_distribution.to_excel(writer, sheet_name='Age Distribution', index=False)
    generation_distribution.to_excel(writer, sheet_name='Generation Distribution', index=False)
    gender_distribution.to_excel(writer, sheet_name='Gender Diversity', index=False)
    tenure_distribution.to_excel(writer, sheet_name='Tenure Analysis', index=False)
    resignation_trends.to_excel(writer, sheet_name='Resignation Trends', index=False)
    retention_cohort.to_excel(writer, sheet_name='Retention by Cohort (Names)', index=False)
    retention_summary.to_excel(writer, sheet_name='Retention by Cohort (Summary)', index=False)
    promotion_transfer_tracking.to_excel(writer, sheet_name='Promotion & Transfer', index=False)
    duplicates.to_excel(writer, sheet_name='Duplicate Names by Cohort', index=False)
    headcount_per_year.to_excel(writer, sheet_name='Headcount Per Year', index=False)
    likert_per_year_percentage.to_excel(writer, sheet_name='Satisfaction Prct', index=False)
    likert_per_year_count.to_excel(writer, sheet_name='Satisfaction Count', index=False)
    engagement_index_per_year.to_excel(writer, sheet_name='Engagement Index', index=False)
    driver_resignation_df.to_excel(writer, sheet_name='Driver-Resignation', index=False)
    driver_promotion_df.to_excel(writer, sheet_name='Driver-Promotion', index=False)
    promotion_predictors_df.to_excel(writer, sheet_name='Promotion Predictors', index=False)
    engagement_retention.to_excel(writer, sheet_name='Engagement vs Retention', index=False)
    satisfaction_retention.to_excel(writer, sheet_name='Satisfaction vs Retention', index=False)

print("Results saved to HR_Analysis_Output.xlsx")
