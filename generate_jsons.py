import json 
from datetime import datetime, timedelta
import os

def generate_json_files(base_json, start_date, num_weeks, output_dir, end_date_delta = timedelta(days=0, hours=13)):
    # Convert the initial date string to a datetime object
    base_date = datetime.strptime(start_date, "%d/%m/%Y %H:%M")
    
    # Loop to generate multiple JSON files
    for i in range(num_weeks):
        # Calculate the dates
        sign_date = base_date - timedelta(days=1)  # Sign date is one day before the start
        end_date = base_date + end_date_delta   # End date is one day after the start

        # Update the base JSON with the new dates
        updated_json = base_json.copy()
        updated_json["[EMPTYDATAS]"] = base_date.strftime("%d/%m/%Y %H:%M")
        updated_json["[EMPTYDATAE]"] = end_date.strftime("%d/%m/%Y %H:%M")
        updated_json["[EMPTYDATASIGN]"] = sign_date.strftime("%d/%m/%Y")

        # Save the JSON to a file
        start_date_print = f"{base_date.day}_{base_date.month}_{base_date.year}" 
        filename = f"{output_dir}/generated_{start_date_print}.json"
        with open(filename, 'w') as json_file:
            json.dump(updated_json, json_file, indent=4)
        
        # Increment the base date for the next file
        base_date += timedelta(days=7)

    print(f"{num_weeks} JSON files generated in '{output_dir}'.")

# Example usage


if __name__=="__main__":
    output_directory = "./missioni_ancona"  # Set your desired output directory
    start_date_input = "25/11/2024 06:30"  # Starting date
    number_of_files = 4  # Number of JSON files to generate
    # Ensure the output directory exists
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Generate JSON files
    base_json = json.load(open("template_ancona.json"))
    generate_json_files(base_json, start_date_input, number_of_files, output_directory)
