# Week 3: Data Integrity and Model Protection

Coursework project on AI model security in cybersecurity. It simulates a data poisoning attack on a phishing email classifier and includes a short research paper on model protection techniques.

## Contents

| File | Description |
| --- | --- |
| `Week_3_Assignment_Data_Integrity_and_Model_Protection.docx` | Full write-up: Part 1 summary, Part 2 paper, APA references |
| `Part_1_Data_Integrity_Analysis_Summary.docx` | Standalone Part 1 summary (about 137 words) |
| `Week 3 mail dataset.csv` | Original 20-email dataset (4 features plus `Spam_Label`) |
| `Week 3 mail dataset - Poisoned Simulation.csv` | Dataset showing the poisoned training samples |
| `Week 3 Phishing Email Dataset and Poisoning Results.xlsx` | Clean data, poisoned data, and evaluation metrics |
| `week3_ai_model_poisoning.py` | Poisoning simulation script |
| `generate_week3_deliverables.py` | Script that rebuilds the CSV, Excel, and Word files |

## Part 1: Data poisoning simulation

A Logistic Regression classifier is trained on email features (subject length, spam keywords, sender reputation, link count). The script then poisons 50% of the training data by flipping labels and adding Gaussian noise (mean 0, standard deviation 20), and retrains.

In the seeded local run (`np.random.seed(42)`), test accuracy fell from 100% to 50%. The test set has only 4 emails, so results vary widely without a fixed seed. The original assignment code sets no seed, so a Colab run may give different numbers.

## Part 2: Model protection techniques

The paper covers differential privacy and homomorphic encryption. It explains how each defends AI models in cybersecurity and how the poisoning exercise relates to offensive AI use.

## Run it

```bash
pip install pandas numpy scikit-learn python-docx openpyxl
python week3_ai_model_poisoning.py
python generate_week3_deliverables.py   # regenerates data files and Word documents
```

The generator script writes to a hardcoded Windows path (`OUTPUT_DIR`). Edit it for your machine.

## Notes

- The title page contains placeholders for the course and instructor name.
- Gemini and Claude were used as AI tools and are cited in the paper.
- This is student coursework. Do not submit it as your own.
