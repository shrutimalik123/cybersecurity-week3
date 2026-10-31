# ---------------------------------------------
# Week 3 Assignment: Phishing Email Model Data Poisoning
# ---------------------------------------------
import pandas as pd
import numpy as np
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score

# Set seed for reproducible simulation results
np.random.seed(42)

# Step 2: Load the Dataset
df = pd.read_csv('Week 3 mail dataset.csv')
print(df.head())

# Step 3: Split Data into Features and Labels
X = df.drop(columns=["Spam_Label"])
y = df["Spam_Label"]

print("Feature Shape:", X.shape)
print("Label Shape:", y.shape)

# Step 4: Split the Dataset into Training and Testing Sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

print("Training set size:", X_train.shape)
print("Test set size:", X_test.shape)

# Step 5: Train a Machine Learning Model on the Original Dataset
model = LogisticRegression()
model.fit(X_train, y_train)
y_pred = model.predict(X_test)

accuracy_original = accuracy_score(y_test, y_pred)
print("Accuracy on Original Dataset:", accuracy_original)

# Step 6: Introduce Data Poisoning (Manipulating the Dataset)
X_train_poisoned = X_train.astype(float).copy()
y_train_poisoned = y_train.copy()

# Poison 50% of the training data for stronger impact
num_samples_to_poison = int(0.5 * len(y_train))
indices_to_flip = np.random.choice(y_train.index, num_samples_to_poison, replace=False)

# Flip the labels (0 <-> 1)
y_train_poisoned.loc[indices_to_flip] = 1 - y_train_poisoned.loc[indices_to_flip]

# Add heavy noise using .loc to avoid index mismatch
noise = np.random.normal(0, 20, X_train_poisoned.loc[indices_to_flip].shape)
X_train_poisoned.loc[indices_to_flip] += noise

print("Number of poisoned samples:", num_samples_to_poison)

# Step 7: Train the Model on the Poisoned Dataset
model_poisoned = LogisticRegression()
model_poisoned.fit(X_train_poisoned, y_train_poisoned)
y_pred_poisoned = model_poisoned.predict(X_test)

accuracy_poisoned = accuracy_score(y_test, y_pred_poisoned)
print("Accuracy on Poisoned Dataset:", accuracy_poisoned)

# Step 8: Compare Results and Analyze Impact
print(f"Original Model Accuracy: {accuracy_original:.4f}")
print(f"Poisoned Model Accuracy: {accuracy_poisoned:.4f}")

accuracy_drop = (accuracy_original - accuracy_poisoned) * 100
print(f"Accuracy Drop Due to Data Poisoning: {accuracy_drop:.2f}%")
