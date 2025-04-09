import pandas as pd
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta

# Load dataset
df = pd.read_excel("dataset.xlsx")

# Exclude specific customers based on name
excluded_customers = [
    "point of sale customer", "samudra", "retail section", "palm tree", 
    "kalyan", "med lab", "calicut healthcare llp"
]
df = df[~df["Order Customer Name"].str.lower().isin(excluded_customers)]

# Convert 'Order Ordered Date' to datetime
df["Order Ordered Date"] = pd.to_datetime(df["Order Ordered Date"], dayfirst=True)

# GUI Function
def find_churned_customers():
    try:
        # Get the selected reference date from input
        reference_date = datetime.strptime(date_entry.get(), "%d-%m-%Y")

        # Compute Recency & Frequency for each customer
        customer_stats = df.groupby("Order Customer Name").agg(
            Last_Purchase_Date=("Order Ordered Date", "max"),
            Contact=("Order Customer Contact", "first"),
            Last_Product=("Product Full Name", "last"),
            Order_Register_Name=("Order Register Name", "last")
        ).reset_index()

        # Calculate recency (days since last purchase)
        customer_stats["Recency"] = (reference_date - customer_stats["Last_Purchase_Date"]).dt.days

        # Create an empty list to store churned customers
        churned_customers_list = []

        # Loop through each customer
        for _, row in customer_stats.iterrows():
            customer_name = row["Order Customer Name"]
            last_purchase_date = row["Last_Purchase_Date"]

            # Condition 1: Latest purchase must be > 180 days from the input date
            if row["Recency"] <= 180:
                continue  # Skip this customer as they are not churned

            # Get all purchase dates for this customer
            customer_purchases = df[df["Order Customer Name"] == customer_name]["Order Ordered Date"].sort_values()

            # Find purchases within 180 days before the latest purchase
            purchases_before_latest = customer_purchases[
                (customer_purchases >= last_purchase_date - timedelta(days=180)) &
                (customer_purchases < last_purchase_date)
            ]

            # Condition 2: Must have at least 3 purchases within 180 days before the latest purchase
            if len(purchases_before_latest) >= 3:
                churned_customers_list.append(row)

        # Convert to DataFrame
        churned_customers = pd.DataFrame(churned_customers_list)

        # Save to CSV if any churned customers are found
        if not churned_customers.empty:
            churned_customers.to_csv("churned_customers.csv", index=False)
            messagebox.showinfo("Success", "Churned customers list saved as 'churned_customers.csv'")
        else:
            messagebox.showinfo("Info", "No churned customers found based on the given criteria.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# GUI Setup
root = tk.Tk()
root.title("Churned Customer Finder")
root.geometry("400x200")

# Label and Entry for Date Input
tk.Label(root, text="Enter Reference Date (DD-MM-YYYY):").pack(pady=5)
date_entry = tk.Entry(root)
date_entry.pack(pady=5)

# Button to Process Data
process_button = tk.Button(root, text="Find Churned Customers", command=find_churned_customers)
process_button.pack(pady=10)

root.mainloop()