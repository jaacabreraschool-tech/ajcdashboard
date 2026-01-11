import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, confusion_matrix, roc_auc_score

# -----------------------------
# Load dataset
# -----------------------------
df = pd.read_excel("HR Cleaned Data 01.09.26.xlsx")

# Extract reporting year
df['Year'] = pd.to_datetime(df['Calendar Year']).dt.year

# -----------------------------
# Prepare target variable (Resignee Checking)
# -----------------------------
df['Resignee_Binary'] = df['Resignee Checking '].apply(
    lambda x: 1 if str(x).upper() in ['LEAVER'] else 0
)

# -----------------------------
# Select predictor variables
# -----------------------------
features = [
    'Tenure', 'Job Satisfaction', 'Pay/Benefits', 'Career',
    'Respect', 'Communication', 'Corporate Culture', 'Management'
]

# Drop rows with missing values in predictors
model_df = df.dropna(subset=features + ['Resignee_Binary'])

X = model_df[features]
y = model_df['Resignee_Binary']

# -----------------------------
# Train/test split
# -----------------------------
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.3, random_state=42, stratify=y
)

# -----------------------------
# Scale features
# -----------------------------
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

# -----------------------------
# Logistic Regression model
# -----------------------------
model = LogisticRegression(max_iter=1000)
model.fit(X_train_scaled, y_train)

# -----------------------------
# Predictions
# -----------------------------
y_pred = model.predict(X_test_scaled)
y_prob = model.predict_proba(X_test_scaled)[:, 1]

# -----------------------------
# Evaluation
# -----------------------------
print("Confusion Matrix:")
print(confusion_matrix(y_test, y_pred))

print("\nClassification Report:")
print(classification_report(y_test, y_pred))

print("\nROC-AUC Score:", roc_auc_score(y_test, y_prob))

# -----------------------------
# Feature importance (coefficients)
# -----------------------------
importance = pd.DataFrame({
    'Feature': features,
    'Coefficient': model.coef_[0]
}).sort_values(by='Coefficient', ascending=False)

print("\nFeature Importance:")
print(importance)

# -----------------------------
# Save predictions back to Excel
# -----------------------------
model_df['Attrition_Risk_Score'] = model.predict_proba(scaler.transform(X))[:, 1]

with pd.ExcelWriter("Attrition_Risk_Output.xlsx") as writer:
    model_df[['Full Name', 'Year', 'Resignee_Binary', 'Attrition_Risk_Score']].to_excel(
        writer, sheet_name="Attrition Risk Scores", index=False
    )
    importance.to_excel(writer, sheet_name="Feature Importance", index=False)

print("Results saved to Attrition_Risk_Output.xlsx")
