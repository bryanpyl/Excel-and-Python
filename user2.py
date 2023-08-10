import pandas as pd
import sys
import os
import PyInstaller.__main__

def generate_user_exe(input_name):
    excel_file = "C:\\Users\\user\\OneDrive\\Desktop\\SAINS\\Week 4\\7.8.2023\\UserList.xlsx"

    try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(excel_file)

        # Check if the required columns exist in the DataFrame
        required_columns = ['Username', 'Password', 'Email']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Error: Required columns not found in Excel file: {', '.join(missing_columns)}")
            return

        # Filter the DataFrame based on the input name
        filtered_df = df[df['Username'] == input_name]

        # Check if the input name exists in the DataFrame
        if filtered_df.empty:
            print("User not found.")
            return

        # Retrieve the Username, Password, and Email for the user
        # df.iloc[row_index, column_index]

        username = filtered_df['Username'].iloc[0]
        password = filtered_df['Password'].iloc[0]
        email = filtered_df['Email'].iloc[0]

        # Prepare the content for the .py file (Python script)
        script_content = f"""
username = "{username}"
password = "{password}"
email = "{email}"

print("Username:", username)
print("Password:", password)
print("Email:", email)
"""

        # Create a temporary .py file
        temp_py_file = f"{username}_details.py"
        with open(temp_py_file, "w") as file:
            file.write(script_content)

        # Use PyInstaller to convert the temporary .py file to .exe
        # Return current working directory and join with subdirectory, user_exe_files
        output_dir = os.path.join(os.getcwd(), "user_exe_files")
        PyInstaller.__main__.run([
            '--onefile',
            '--distpath',
            output_dir,
            temp_py_file
        ])

        # Clean up the temporary .py file
        os.remove(temp_py_file)

        print(f"{username}.exe generated with user details.")

    except FileNotFoundError:
        print("Excel file not found.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <input_name>")
    else:
        input_name = sys.argv[1]
        generate_user_exe(input_name)
