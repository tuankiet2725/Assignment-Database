import pandas as pd
from tkinter import Tk, filedialog

# Truck capacities
max_weight = 500
max_volume = 2000

distance_matrix = {
    "Hanoi": {"HCMC": 1700, "Da Nang": 800, "Nha Trang": 1300, "Hai Phong": 100, "Dalat": 1400},
    "HCMC": {"Hanoi": 1700, "Da Nang": 850, "Nha Trang": 400, "Hai Phong": 1600, "Dalat": 300},
    "Da Nang": {"Hanoi": 800, "HCMC": 850, "Nha Trang": 550, "Hai Phong": 900, "Dalat": 600},
    "Nha Trang": {"Hanoi": 1300, "HCMC": 400, "Da Nang": 550, "Hai Phong": 1250, "Dalat": 250},
    "Hai Phong": {"Hanoi": 100, "HCMC": 1600, "Da Nang": 900, "Nha Trang": 1250, "Dalat": 1350},
    "Dalat": {"Hanoi": 1400, "HCMC": 300, "Da Nang": 600, "Nha Trang": 250, "Hai Phong": 1350},
}

# Ensure the distance from Hanoi to itself exists
distance_matrix["Hanoi"]["Hanoi"] = 0


# Function to read data from an Excel file
def read_excel_file():
    try:
        root = Tk()
        root.withdraw()  # Hide the Tkinter root window
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            print("No file selected.")
            return None

        df = pd.read_excel(file_path)
        if all(col in df.columns for col in ["Destination", "Weight", "Volume"]):
            return df.to_dict(orient="records")
        else:
            raise ValueError("The Excel file must have 'Destination', 'Weight', and 'Volume' columns.")
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

# Function to select parcels based on truck capacity
def select_parcels(goods, max_weight, max_volume):
    selected_parcels = []
    current_weight, current_volume = 0, 0
    goods.sort(key=lambda x: x["Weight"] / (x["Volume"] or 1), reverse=True)
    for item in goods:
        if current_weight + item["Weight"] <= max_weight and current_volume + item["Volume"] <= max_volume:
            selected_parcels.append(item)
            current_weight += item["Weight"]
            current_volume += item["Volume"]
    return selected_parcels

# Function to calculate the nearest neighbor TSP route
def tsp_nearest_neighbor(cities):
    current_city = "Hanoi"
    unvisited = set(cities[1:])
    route = [current_city]
    total_distance = 0

    while unvisited:
        next_city = min(unvisited, key=lambda city: distance_matrix[current_city][city])
        total_distance += distance_matrix[current_city][next_city]
        current_city = next_city
        route.append(current_city)
        unvisited.remove(current_city)

    # Return to the starting city
    total_distance += distance_matrix[current_city]["Hanoi"]
    route.append("Hanoi")
    return route, total_distance

# Main function
def main():
    goods = read_excel_file()
    if not goods:
        return

    selected_parcels = select_parcels(goods, max_weight, max_volume)
    if not selected_parcels:
        print("No parcels could be selected within truck capacity.")
        return

    print("\nSelected Parcels:")
    for parcel in selected_parcels:
        print(f"{parcel['Destination']}: Weight={parcel['Weight']}kg, Volume={parcel['Volume']}m³")

    destinations = ["Hanoi"] + [parcel["Destination"] for parcel in selected_parcels]
    optimal_route, total_distance = tsp_nearest_neighbor(destinations)

    print("\nOptimal Route:")
    for i, city in enumerate(optimal_route[:-1]):
        print(f"{i + 1}. {city}")
    print(f"\nTotal Distance: {total_distance} km")

if __name__ == "__main__":
    main()

