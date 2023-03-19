import email
import imaplib
import re

import pandas as pd


def select_mailbox(email_address, password, mailbox_name):
    """
    This function selects the IMAP mailbox with the given name for the given email address and password
    """
    # create an IMAP client instance
    mail = imaplib.IMAP4_SSL("imap.gmail.com")

    # log in to the Gmail account
    mail.login(email_address, password)

    # select the mailbox with the given name
    mail.select(mailbox_name)

    return mail


def search_emails_by_keyword(mail, keyword_to_search, emails_to_search):
    """
    This function searches the latest N email messages (where N is specified by the user)
    in the selected mailbox for the given keyword in the message body,
    and returns a set of unique sender email addresses that match the search criteria
    """
    # get the latest N email messages from the mailbox
    typ, data = mail.search(None, "ALL")
    latest_email_ids = data[0].split()[-emails_to_search:]

    print(
        f'Searching for keyword "{keyword_to_search}" in the body of the latest '
        + str(emails_to_search)
        + " email messages...\n"
    )

    count_keyword = 0
    sender_set = (
        set()
    )  # create an empty set to store unique email addresses of senders

    # iterate over the latest email messages
    for email_id in latest_email_ids:
        typ, msg_data = mail.fetch(email_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                # parse the email message
                msg = email.message_from_bytes(response_part[1])

                # get the sender of the email
                sender = msg["From"]

                # extract the body of the email message as text
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            try:
                                body = part.get_payload(decode=True).decode(
                                    "utf-8"
                                )
                            except UnicodeDecodeError:
                                body = part.get_payload(decode=True).decode(
                                    "iso-8859-1"
                                )

                            count_keyword += body.count(keyword_to_search)

                            # add the email address to the set if the keyword is found
                            if keyword_to_search.lower() in body.lower():
                                email_address = re.search(
                                    r"[\w\.-]+@[\w\.-]+", sender
                                )
                                sender_set.add(email_address.group(0))

                        elif part.get_content_maintype() == "text":
                            try:
                                body = part.get_payload(decode=True).decode(
                                    "utf-8"
                                )
                            except UnicodeDecodeError:
                                body = part.get_payload(decode=True).decode(
                                    "iso-8859-1"
                                )

                            count_keyword += body.count(keyword_to_search)

                            # add the email address to the set if the keyword is found
                            if keyword_to_search.lower() in body.lower():
                                email_address = re.search(
                                    r"[\w\.-]+@[\w\.-]+", sender
                                )
                                sender_set.add(email_address.group(0))
                else:
                    try:
                        body = msg.get_payload(decode=True).decode("utf-8")
                    except UnicodeDecodeError:
                        body = msg.get_payload(decode=True).decode(
                            "iso-8859-1"
                        )

                    count_keyword += body.count(keyword_to_search)

                    # add the email address to the set if the keyword is found
                    if keyword_to_search.lower() in body.lower():
                        email_address = re.search(r"[\w\.-]+@[\w\.-]+", sender)
                        sender_set.add(email_address.group(0))

    # print the total number of occurrences of the keyword
    print(
        f'The keyword "{keyword_to_search}" appears {count_keyword} time(s) in the body of the latest '
        + str(emails_to_search)
        + " email messages. \n"
    )

    return sender_set, count_keyword


def export_sender_emails_to_excel(sender_set, filepath="sender_emails.xlsx"):
    """
    This function creates a pandas DataFrame with the given set of
    unique sender email addresses, and exports the DataFrame to an Excel file
    named "sender_emails.xlsx" in the same directory as this Python script.
    """
    # create pandas DataFrame with sender_set data
    df = pd.DataFrame(list(sender_set), columns=["Sender Email"])

    # export DataFrame to Excel file
    df.to_excel(filepath, index=False)

    print(
        f'Successfully exported {len(sender_set)} unique sender email addresses to "{filepath}" file.\n'
    )
    print(
        "your excel file is stored in the same location as this program. \nDownload it and use Excel to open the file."
    )


def generate_excel_file(
    email_address,
    password,
    keyword_to_search,
    emails_to_search,
    filepath="sender_emails.xlsx",
):
    # select the Gmail inbox mailbox
    mail = select_mailbox(email_address, password, "inbox")

    # search for unique sender email addresses by keyword
    sender_set, count_keyword = search_emails_by_keyword(
        mail, keyword_to_search, emails_to_search
    )

    # export the unique sender email addresses to an Excel file
    export_sender_emails_to_excel(sender_set, filepath)


if __name__ == "__main__":
    # get user input for email search parameters
    email_address = input("what is your email address?: \n")
    password = input(
        "what is your App password? (This is a special password under security under settings): \n"
    )
    keyword_to_search = input("what keyword do you want to search?: \n")
    emails_to_search = int(
        input(
            "how many emails do you want to search through? (integer input): \n"
        )
    )

    generate_excel_file(
        email_address, password, keyword_to_search, emails_to_search
    )
