import os

# List of branch names
branches = [
    'Bacolod Branch',
    'Batac Branch',
    'Butuan Branch',
    'Cabanatuan Branch',
    'Cagayan De Oro Branch',
    'Dagupan Branch',
    'Dumaguete Branch',
    'General Santos Branch',
    'Head Office',
    'Iloilo Branch',
    'La Union Branch',
    'Legazpi Branch',
    'Lucena Branch',
    'Mindanao Regional Office',
    'North Luzon Regional Office',
    'Ozamiz Branch',
    'Puerto Princesa Branch',
    'Roxas Branch',
    'Tacloban Branch',
    'Tuguegarao Branch',
    'Visayas Regional Office',
    'Zamboanga Branch',
]

# Target directory
base_dir = "/Users/finellajianna/carbon-emissions/Data"

# Create folders for each branch
for name in branches:
    path = os.path.join(base_dir, name)
    os.makedirs(path, exist_ok=True)
    print(f"Created: {path}")

print("All folders created successfully!")
