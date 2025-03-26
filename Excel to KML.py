# Import the necessary libraries
import pandas as pd
import simplekml

# Step 1: Load the Excel file
file_path = r"C:\Users\CHAKU FOODS\Documents\Data Collection Cleaning\Farm Table_Updated.xlsx" #use the file path on your own local disk
data = pd.read_excel(file_path)

data.head()

# Initialize the KML object
kml = simplekml.Kml()

# Step 2: Function to parse coordinates and generate KML file
def parse_coordinates(coordinates_str):
    geopins = []
    if isinstance(coordinates_str, str):  # Ensure the input is a string
        for coord in coordinates_str.split(','):  # Split coordinates by commas
            parts = coord.strip().split()  # Split each coordinate into parts (latitude, longitude, etc.)
            if len(parts) >= 2:  # Ensure we have at least two parts for lat and lon
                try:
                    latitude = float(parts[0])
                    longitude = float(parts[1])
                    geopins.append((longitude, latitude))  # KML format is (lon, lat)
                except ValueError:
                    print(f"Skipping invalid coordinate: {coord}")
    return geopins

# Step 3: Create a new KML object
kml = simplekml.Kml()

# Define a style for the polygons (transparent color)
polystyle = simplekml.Style()
polystyle.polystyle.color = '7d00ff00'  # 7d is the transparency, 00ff00 is green color

# Step 4: Loop through each row in the DataFrame and add polygons to the KML
for index, row in data.iterrows():
    farmer_name = f"{row['last_name']} {row['first_name']}"  # Concatenate last_name and first_name to get the full name
    geopins_raw = row['geo_boundaries']  # Get raw coordinates from the 'geo_boundaries' column
    
    # Parse the geographic boundaries to extract coordinates
    coordinates = parse_coordinates(geopins_raw)
    
    if coordinates:
        # Create a polygon for the farm
        pol = kml.newpolygon(name=farmer_name)
        pol.outerboundaryis = coordinates
        pol.style = polystyle

# Step 5: Save the KML to a file
kml.save("Test.kml")
print("KML file created successfully!")
