import subprocess

# Define paths to the scripts
generate_script = 'NMAP.py'
compare_script = 'compare.py'

# Step 1: Run the script to generate the new Excel file
print("Running generate_new_excel.py...")
generate_process = subprocess.run(['python', generate_script])

# Check if the generation script completed successfully
if generate_process.returncode == 0:
    print("Generation script completed successfully. Proceeding to comparison.")
    
    # Step 2: Run the script to compare with the old file and generate unique entries
    compare_process = subprocess.run(['python', compare_script])
    
    # Check if the comparison script completed successfully
    if compare_process.returncode == 0:
        print("Comparison script completed successfully.")
    else:
        print("Comparison script encountered an error.")
else:
    print("Generation script encountered an error. Comparison will not run.")
