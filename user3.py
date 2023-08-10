import pandas as pd

def retrieve_info_from_excel(excel_file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(excel_file_path)

        # Assuming the columns containing username, password, and email are named 'Username', 'Password', and 'Email'
        usernames = df['Username'].tolist()
        passwords = df['Password'].tolist()
        emails = df['Email'].tolist()

        # Return the retrieved data
        return df, usernames, passwords, emails

    except Exception as e:
        print(f"Error occurred: {e}")
        return None, [], [], []

def save_info_to_excel(df, excel_file_path):
    try:
        # Save the DataFrame to the Excel file
        df.to_excel(excel_file_path, index=False)
        print("Data saved to the Excel file successfully.")
    except Exception as e:
        print(f"Error occurred while saving data: {e}")

def add_new_data(df):
    username = input("Enter new username: ")
    password = input("Enter new password: ")
    email = input("Enter new email: ")

    new_data = pd.DataFrame({'Username': [username], 'Password': [password], 'Email': [email]})
    # Concatenate original with new DataFrame, ensure new DataFrame has new index
    df = pd.concat([df, new_data], ignore_index=True)
    return df

def edit_data(df):
    username_to_edit = input("Enter the username to edit: ")

    if username_to_edit not in df['Username'].values:
        print("Username not found.")
        return df

    new_password = input("Enter the new password: ")
    new_email = input("Enter the new email: ")

    # Access or modify data based on label and condition
    # df.loc[row_labels, column_labels]
    df.loc[df['Username'] == username_to_edit, 'Password'] = new_password
    df.loc[df['Username'] == username_to_edit, 'Email'] = new_email

    return df

def delete_data(df):
    username_to_delete = input("Enter the username to delete: ")

    if username_to_delete not in df['Username'].values:
        print("Username not found.")
        return df

    df = df[df['Username'] != username_to_delete]
    return df

def display_user_info(df, username):
    user_info = df[df['Username'] == username]
    if not user_info.empty:
        return user_info
    else:
        print("Username not found.")
        return None

if __name__ == "__main__":
    # Define the path to the Excel file (embedded in the script)
    excel_file_path = "C:\\Users\\user\\OneDrive\\Desktop\\SAINS\\Week 4\\7.8.2023\\UserList.xlsx"  # Replace with your actual Excel file path

    # Call the function to retrieve information from the Excel file
    df, _, _, _ = retrieve_info_from_excel(excel_file_path)

    # Check if the DataFrame was successfully loaded
    if df is not None:
        while True:
            print("\nWhat would you like to do?")
            print("1. Add new data")
            print("2. Edit existing data")
            print("3. Delete data")
            print("4. Display user information")
            print("5. Save and exit")
            choice = input("Enter the number of your choice (1/2/3/4/5): ")

            if choice == "1":
                df = add_new_data(df)
            elif choice == "2":
                df = edit_data(df)
            elif choice == "3":
                df = delete_data(df)
            elif choice == "4":
                username_to_display = input("Enter the username to display information: ")
                user_info = display_user_info(df, username_to_display)
                if user_info is not None:
                    print("\nUser Information:")
                    print(user_info)
            elif choice == "5":
                save_info_to_excel(df, excel_file_path)
                break
            else:
                print("Invalid choice. Please enter a valid number.")

    else:
        print("No data found or an error occurred while retrieving information.")
