import pandas as pd
from sklearn.naive_bayes import GaussianNB
from sklearn.preprocessing import StandardScaler
import tkinter as tk
from tkinter import ttk, messagebox
import os

# Excel path
excel_file = r"C:\Users\ganesan\Music\mini project\smart_meter_data.xlsx"
os.makedirs(os.path.dirname(excel_file), exist_ok=True)

# Function to create default labeled Excel if missing or empty
def create_default_excel():
    df = pd.DataFrame({
        'Voltage': [220]*12,
        'Current': [0.3,0.5,1.2,1.4,3.0,3.3,0.08,0.09,3.6,3.9,0.12,0.13],
        'Power':   [66,110,264,308,660,726,17.6,19.8,792,858,26.4,28.6],
        'Appliance': ['Fan','Fan','Fridge','Fridge','AC','AC','Light','Light','Heater','Heater','LED','LED']
    })
    df.to_excel(excel_file, index=False)
    print("✅ Excel created with default labeled rows.")

# Check if Excel exists and has labeled data
if not os.path.exists(excel_file):
    create_default_excel()

df = pd.read_excel(excel_file, engine='openpyxl')
if 'Appliance' not in df.columns or df['Appliance'].isna().all():
    create_default_excel()
    df = pd.read_excel(excel_file, engine='openpyxl')

df['Appliance'] = df['Appliance'].astype(object)

# Train model
def train_model():
    labeled_df = df[df['Appliance'].notna()]
    X_train = labeled_df[['Current','Power']].values
    y_train = labeled_df['Appliance']
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X_train)
    model = GaussianNB()
    model.fit(X_scaled, y_train)
    return model, scaler    

model, scaler = train_model()

# GUI setup
root = tk.Tk()
root.title("Smart Meter Appliance Predictor")
root.geometry("700x550")

columns = ['Voltage','Current','Power','Appliance']
tree = ttk.Treeview(root, columns=columns, show='headings')
for col in columns:
    tree.heading(col, text=col)
tree.pack(fill=tk.BOTH, expand=True)

# Insert existing Excel rows
for _, row in df.iterrows():
    tree.insert('', tk.END, values=(row['Voltage'], row['Current'], row['Power'], row['Appliance']))

# Input fields
input_frame = tk.Frame(root)
input_frame.pack(pady=10)
tk.Label(input_frame, text="Voltage (V):").grid(row=0, column=0)
voltage_entry = tk.Entry(input_frame)
voltage_entry.grid(row=0, column=1)
tk.Label(input_frame, text="Current (A):").grid(row=1, column=0)
current_entry = tk.Entry(input_frame)
current_entry.grid(row=1, column=1)

result_label = tk.Label(root, text="", font=("Arial", 14))
result_label.pack(pady=10)

# Predict button function
def predict_appliance():
    global df, model, scaler
    try:
        voltage = float(voltage_entry.get())
        current = float(current_entry.get())
        power = voltage * current

        X_new_scaled = scaler.transform([[current, power]])
        predicted = model.predict(X_new_scaled)[0]

        result_label.config(text=f"✅ Predicted Appliance: {predicted}", fg="green")

        new_row = {'Voltage': voltage, 'Current': current, 'Power': power, 'Appliance': predicted}
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(excel_file, index=False)
        tree.insert('', tk.END, values=(voltage, current, power, predicted))

        # Retrain model with new data
        model, scaler = train_model()

        voltage_entry.delete(0, tk.END)
        current_entry.delete(0, tk.END)

    except ValueError:
        messagebox.showerror("Error", "Please enter valid numbers!")

tk.Button(root, text="Predict Appliance", command=predict_appliance, bg="green", fg="white").pack(pady=10)

root.mainloop()
