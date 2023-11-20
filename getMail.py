from exchangelib import Credentials, Account, DELEGATE

def get_outlook_email(first_name, last_name):
    # Replace with your Outlook credentials
    username = '***********'
    password = '*****'

    # Connect to Outlook
    credentials = Credentials(username, password)
    account = Account(username, credentials=credentials, autodiscover=True, access_type=DELEGATE)

    # Search for the user based on first and last name
    resolved_names, status = account.protocol.resolve_names([f"{first_name} {last_name}"],
                                                            return_full_contact_data=True)

    if resolved_names:
        # Get the first resolved name (assuming the list is not empty)
        first_resolved_name = resolved_names[0]

        # Get the user's email address
        email_address = first_resolved_name.email_address
        return email_address
    else:
        print(f"User {first_name} {last_name} not found in Outlook directory.")
        return None

if __name__ == "__main__":
    first_name = "Sachin"
    last_name = "Balyan"


    email_address = get_outlook_email(first_name, last_name)

    if email_address:
        print(f"Email address for {first_name} {last_name}: {email_address}")
    else:
        print("Email address not found.")
