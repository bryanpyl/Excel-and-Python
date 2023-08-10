import pandas as pd
import sys

def retrieve_user_info(input_name):
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
        username = filtered_df['Username'].iloc[0]
        password = filtered_df['Password'].iloc[0]
        email = filtered_df['Email'].iloc[0]

        # Display the retrieved information
        print(f"Username: {username}")
        print(f"Password: {password}")
        print(f"Email: {email}")

    except FileNotFoundError:
        print("Excel file not found.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <input_name>")
    else:
        input_name = sys.argv[1]
        retrieve_user_info(input_name)
