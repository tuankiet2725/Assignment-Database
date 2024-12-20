import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import itertools

# Distance matrix (in km) between cities
distance_matrix = {
    "Hanoi": {"HCMC": 1700, "Da Nang": 800, "Nha Trang": 1300, "Hai Phong": 100, "Dalat": 1400},
    "HCMC": {"Hanoi": 1700, "Da Nang": 850, "Nha Trang": 400, "Hai Phong": 1600, "Dalat": 300},
    "Da Nang": {"Hanoi": 800, "HCMC": 850, "Nha Trang": 550, "Hai Phong": 900, "Dalat": 600},
    "Nha Trang": {"Hanoi": 1300, "HCMC": 400, "Da Nang": 550, "Hai Phong": 1250, "Dalat": 250},
    "Hai Phong": {"Hanoi": 100, "HCMC": 1600, "Da Nang": 900, "Nha Trang": 1250, "Dalat": 1350},
    "Dalat": {"Hanoi": 1400, "HCMC": 300, "Da Nang": 600, "Nha Trang": 250, "Hai Phong": 1350},
}

# Truck capacities (weight and volume)
max_weight = 500  # Maximum weight the truck can carry (kg)
max_volume = 2000  # Maximum volume the truck can carry (cubic meters)

# Cost rates
rate_per_kg_km = 50
rate_per_m3_km = 20
rate_per_km = 100

# Main Application Class
class TruckDeliveryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hanoi Roadways - Truck Delivery System")
        self.parcels = None  # Placeholder for parcel data
        self.create_widgets()

    def create_widgets(self):
        # Instructions and Load Button
        self.instructions_label = tk.Label(self.root, text="Load parcel data from an Excel file.")
        self.instructions_label.pack(pady=10)

        self.load_file_button = tk.Button(self.root, text="Load Data", command=self.load_data)
        self.load_file_button.pack(pady=10)

        # Buttons for actions
        self.calculate_button = tk.Button(self.root, text="Calculate Delivery Route", command=self.calculate_route)
        self.calculate_button.pack(pady=10)

        self.generate_invoice_button = tk.Button(self.root, text="Generate Invoice", command=self.generate_invoice)
        self.generate_invoice_button.pack(pady=10)

        # Frame for results
        self.result_frame = tk.Frame(self.root)
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Placeholder for the table
        self.invoice_frame = None


#LOAD DATA
    def load_data(self):
        try:
            # Open a file dialog to select the Excel file
            file_path = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[("Excel Files", "*.xlsx *.xls")]
            )
            if not file_path:
                return  # User canceled the dialog

            self.parcels = pd.read_excel(file_path)
            required_columns = {"Parcel ID", "Name", "Weight (kg)", "Volume (m³)", "Destination"}
            if not required_columns.issubset(self.parcels.columns):
                raise ValueError(f"The file is missing required columns: {required_columns}.")

            messagebox.showinfo("Success", "Data loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            self.parcels = None

    def calculate_route(self):
        if self.parcels is None:
            messagebox.showerror("Error", "No data loaded. Please load the Excel file first.")
            return

        # Filter parcels within the truck's capacity
        selected_parcels = self.select_parcels(self.parcels)
        if selected_parcels.empty:
            messagebox.showinfo("Info", "No valid parcels selected within the truck's capacity.")
            return

        # Extract unique destinations
        unique_destinations = ["Hanoi"] + selected_parcels["Destination"].unique().tolist()

        # Solve TSP to find the optimal delivery route
        optimal_route, total_distance = self.tsp_nearest_neighbor(unique_destinations)
        result_text = f"Optimal Delivery Route: {' → '.join(optimal_route)}\nTotal Distance: {total_distance} km"
        self.show_result(result_text)

    def select_parcels(self, parcels):
        current_weight = 0
        current_volume = 0

        # Sort parcels by weight-to-volume ratio (descending)
        parcels["Weight_to_Volume"] = parcels["Weight (kg)"] / parcels["Volume (m³)"]
        parcels = parcels.sort_values(by="Weight_to_Volume", ascending=False)

        selected_parcels = []
        for _, parcel in parcels.iterrows():
            if current_weight + parcel["Weight (kg)"] <= max_weight and current_volume + parcel["Volume (m³)"] <= max_volume:
                selected_parcels.append(parcel)
                current_weight += parcel["Weight (kg)"]
                current_volume += parcel["Volume (m³)"]

        return pd.DataFrame(selected_parcels)

    def tsp_nearest_neighbor(self, cities):
        current_city = "Hanoi"
        unvisited = set(cities[1:])  # Exclude the starting city
        route = [current_city]
        total_distance = 0

        while unvisited:
            next_city = min(unvisited, key=lambda city: distance_matrix[current_city][city])
            total_distance += distance_matrix[current_city][next_city]
            current_city = next_city
            route.append(current_city)
            unvisited.remove(current_city)

        # Return to starting city
        total_distance += distance_matrix[current_city]["Hanoi"]
        route.append("Hanoi")
        return route, total_distance

    def generate_invoice(self):
        if self.parcels is None:
            messagebox.showerror("Error", "No data loaded. Please load the Excel file first.")
            return

        invoices = self.generate_invoice_data(self.parcels)

        # Create a scrollable table
        self.clear_previous_results()
        canvas = tk.Canvas(self.result_frame)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="center")
        canvas.pack(side="top", fill="both", expand=True)
        

        # Create table headers
        columns = ["Parcel ID", "Name", "Weight (kg)", "Volume (m³)", "Destination", "Weight Cost (VND)", "Volume Cost (VND)", "Distance Cost (VND)", "Total Cost (VND)"]
        for col_num, col_name in enumerate(columns):
            tk.Label(scrollable_frame, text=col_name, borderwidth=1, relief="solid", width=20, anchor="center").grid(row=0, column=col_num)

        # Populate table with invoice data
        for row_num, invoice in invoices.iterrows():
            for col_num, col_name in enumerate(columns):
                value = invoice[col_name]
                tk.Label(scrollable_frame, text=value, borderwidth=1, relief="solid", width=20).grid(row=row_num + 1, column=col_num)

    def generate_invoice_data(self, parcels):
        invoices = []
        for _, parcel in parcels.iterrows():
            city = parcel["Destination"]
            weight = parcel["Weight (kg)"]
            volume = parcel["Volume (m³)"]
            distance = distance_matrix["Hanoi"].get(city, 0)

            # Calculate costs
            weight_cost = weight * distance * rate_per_kg_km
            volume_cost = volume * distance * rate_per_m3_km
            distance_cost = distance * rate_per_km
            total_cost = weight_cost + volume_cost + distance_cost

            invoices.append({
                "Parcel ID": parcel["Parcel ID"],
                "Name": parcel["Name"],
                "Weight (kg)": weight,
                "Volume (m³)": volume,
                "Destination": city,
                "Weight Cost (VND)": round(weight_cost, 2),
                "Volume Cost (VND)": round(volume_cost, 2),
                "Distance Cost (VND)": round(distance_cost, 2),
                "Total Cost (VND)": round(total_cost, 2),
            })
        return pd.DataFrame(invoices)

    def show_result(self, result_text):
        self.clear_previous_results()  # Clear old results
        result_label = tk.Label(self.result_frame, text=result_text, wraplength=600)
        result_label.pack(pady=50)

    def clear_previous_results(self):
        for widget in self.result_frame.winfo_children():
            widget.destroy()

# Create the Tkinter window and run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = TruckDeliveryApp(root)
    root.mainloop()
