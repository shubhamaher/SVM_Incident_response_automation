import pandas as pd
from sklearn import svm
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
from sklearn.model_selection import cross_val_score
from sklearn.preprocessing import MaxAbsScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
import win32com.client as win32
import os
import matplotlib.pyplot as plt
import datetime

# Load the datasets
train_data = pd.read_csv('UNSW_NB15_training-set.csv')
test_data = pd.read_csv('UNSW_NB15_testing-set.csv')

# Separate the features and labels
X_train = train_data.drop('attack_cat', axis=1)
y_train = train_data['attack_cat']
X_test = test_data.drop('attack_cat', axis=1)
y_test = test_data['attack_cat']

# Identify categorical features automatically
categorical_features = X_train.select_dtypes(include=['object']).columns

# Preprocess categorical features
ct = ColumnTransformer([('encoder', OneHotEncoder(handle_unknown='ignore'), categorical_features)], remainder='passthrough')
X_train = ct.fit_transform(X_train)
X_test = ct.transform(X_test)

# Preprocess numerical features
scaler = MaxAbsScaler()
X_train = scaler.fit_transform(X_train)
X_test = scaler.transform(X_test)

# Train the SVM model
svm_model = svm.SVC(kernel="linear", C=1.0)

# Perform cross-validation
cv_scores = cross_val_score(svm_model, X_train, y_train, cv=5)

# Print the cross-validation scores
print("Cross-Validation Scores:")
for score in cv_scores:
    print(score)

# Fit the model on the full training set
svm_model.fit(X_train, y_train)

# Make predictions on the test set
y_pred = svm_model.predict(X_test)

# Calculate accuracy
accuracy = accuracy_score(y_test, y_pred)
print(f"Accuracy: {accuracy}")

# Check if accuracy is within the desired range
if accuracy >= 0.80 and accuracy <= 0.90:
    print("Accuracy is within the desired range.")
else:
    print("Accuracy is not within the desired range.")

# Evaluate the SVM model
print("SVM Classification Report:")
print(classification_report(y_test, y_pred))
print("SVM Confusion Matrix:")
print(confusion_matrix(y_test, y_pred))


# Create a DataFrame with the test data and predictions
result_df = pd.DataFrame({'Test Data': test_data.values.tolist(), 'Predicted Labels': y_pred})

# Get the current timestamp
timestamp = datetime.datetime.now()

# Add the timestamp column to the DataFrame
result_df['Timestamp'] = timestamp

# Save the DataFrame to an Excel file
result_df.to_excel('SVM_Predictions_Report.xlsx', index=False)

# Create two DataFrames, one for normal attacks and one for all other predicted attacks.
normal_attacks_df = result_df[result_df['Predicted Labels'] == 'Normal']
other_attacks_df = result_df[result_df['Predicted Labels'] != 'Normal']

# Save each DataFrame to a separate Excel file.
normal_attacks_df.to_excel('normal_attacks.xlsx', index=False)
other_attacks_df.to_excel('high_severity_attacks.xlsx', index=False)

# Plot detected attack categories using a pie chart
detected_attacks_counts = result_df['Predicted Labels'].value_counts()
labels = detected_attacks_counts.index
counts = detected_attacks_counts.values

plt.figure(figsize=(10, 6))
plt.pie(counts, labels=labels, autopct=lambda pct: f"{pct:.0f} ({int(round(pct * sum(counts) / 100))})", startangle=90)
plt.title("Detected Attack Categories")
plt.savefig("Detected_Attacks_Pie_Chart.png")
plt.close()

# Function to send an email with attachments using Outlook
def send_email_with_attachments(subject, body, recipient_email, attachments):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body

    for attachment_path, attachment_name in attachments:
        attachment_abs_path = os.path.abspath(attachment_path)
        mail.Attachments.Add(attachment_abs_path, DisplayName=attachment_name)

    mail.Recipients.Add(recipient_email)
    mail.Send()

# Attach the pie chart and the report
subject = 'Attack Report'
body = 'Please find the attack report attached.'
attachments = [
    ('Detected_Attacks_Pie_Chart.png', 'Detected_Attacks_Pie_Chart.png'),
    ('SVM_Predictions_Report.xlsx', 'SVM_Predictions_Report.xlsx')
]

recipient_email = 'your@mail'
send_email_with_attachments(subject, body, recipient_email, attachments)

# Send email for normal attacks
subject_normal = 'Normal Attack Report'
body_normal = 'Please find the normal attack report attached.'
attachments_normal = [('normal_attacks.xlsx', 'normal_attacks.xlsx')]
recipient_email_normal = 'your@mail'
send_email_with_attachments(subject_normal, body_normal, recipient_email_normal, attachments_normal)

# Send email for other predicted attacks
subject_other = 'Other Predicted Attack Report'
body_other = 'Please find the other predicted attack report attached.'
attachments_other = [('high_severity_attacks.xlsx', 'high_severity_attacks.xlsx')]
recipient_email_other = 'your@mail'
send_email_with_attachments(subject_other, body_other, recipient_email_other, attachments_other)
