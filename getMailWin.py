import win32com.client


def get_outlook_email(first_name, last_name):
    outlook_app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get the Contacts folder
    contacts_folder = outlook_app.GetNamespace("MAPI").GetDefaultFolder(10)  # 10 corresponds to the Contacts folder

    # Search for the user based on first and last name
    filter_str = f"[FirstName]='{first_name}' AND [LastName]='{last_name}'"
    user = contacts_folder.Items.Find(filter_str)

    if user:
        # Get the user's email address
        email_address = user.Email1Address
        return email_address
    else:
        print(f"User {first_name} {last_name} not found in Outlook contacts.")
        return None


if __name__ == "__main__":
    first_name = "Sachin"
    last_name = "Balyan"

    email_address = get_outlook_email(first_name, last_name)

    if email_address:
        print(f"Email address for {first_name} {last_name}: {email_address}")
    else:
        print("Email address not found.")
