"""
@Author: 60082363
Script to interact with Outlook app via MAPI using win32com.
TENARIS SOLUTIONS 2023.
"""
import argparse
import win32com.client
import os
import re
import pandas as pd
from datetime import datetime, timedelta


class Outlook:
    def __init__(self, o_mailbox, o_folder) -> None:
        self.o_mailbox = o_mailbox
        self.o_folder = o_folder

        # Creating an object for the outlook application.
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")

        # Creating an object to access outlook folder
        if self.o_mailbox:
            self.inbox = self.ns.Folders[self.o_mailbox].Folders[self.o_folder]
        else:
            self.inbox = self.ns.GetDefaultFolder(6)
        # Update inbox
        self.inbox.GetTable()
        # Getting folder email items
        self.messages = self.inbox.Items

    # Auxiliar functions
    def close(self):
        # close the MAPI object
        self.outlook.Application.Quit()

    def clean_string(self, string):
        string = str(
            string.replace("\n", " ").replace("\t", " ").replace("\r", " ")
        )

        return string

    def format_email_add(self, email_add):
        address = self.ns.CreateRecipient(email_add)
        address.Resolve()
        address_entry = address.AddressEntry
        exchange_user = address_entry.GetExchangeUser()
        email_add = exchange_user.PrimarySmtpAddress

        return email_add

    # End of auxiliar functions

    # Get email items
    def get_emails(self, o_filter, folder_path):
        self.o_filter = o_filter

        # Filter email
        filteredEmails = self.messages.Restrict(self.o_filter)
        # Creating an object to access items inside the inbox of outlook.
        self.messages = filteredEmails

        # Columns for dataframe
        df_columns = [
            "id",
            "store_id",
            "receiver",
            "cc",
            "subject",
            "body",
            "attachments",
            "attachments_count",
            "received",
            "sent",
            "sender",
            "sender_add",
            "unread",
            "html_body",
        ]

        # To iterate through inbox emails using inbox.Items object.
        df_rows = []
        for message in self.messages:
            id = message.EntryID
            store_id = message.Parent.StoreID
            recipients = message.Recipients
            receiver = []
            cc = []
            for recipient in recipients:
                address = recipient.AddressEntry.Address
                try:
                    address = self.format_email_add(address)
                except:
                    pass
                if recipient.Type == 1:
                    receiver.append(address)
                elif recipient.Type == 2:
                    cc.append(address)
            subject = message.Subject
            body = self.clean_string(message.Body)
            html_body = self.clean_string(message.HTMLBody)
            attachments_raw = message.Attachments
            attachments = [att.FileName for att in attachments_raw]
            attachments_count = len(attachments)
            received = message.ReceivedTime.strftime("%m/%d/%y %H:%M:%S")
            sent = message.SentOn.strftime("%m/%d/%y %H:%M:%S")
            sender = message.SenderName
            sender_add = message.SenderEmailAddress
            # Start format email
            try:
                sender_add = self.format_email_add(sender_add)
            except:
                pass
            # End format sender email
            unread = message.UnRead

            new_row = [
                id,
                store_id,
                ";".join(receiver),
                ";".join(cc),
                subject,
                body,
                "|".join(attachments),
                attachments_count,
                received,
                sent,
                sender,
                sender_add,
                unread,
                html_body,
            ]

            # Check if any empty value
            verified_row = ["empty" if x == "" else x for x in new_row]
            df_rows.append(verified_row)

        df = pd.DataFrame(df_rows, columns=df_columns)
        df.to_excel(f"{folder_path}\\df.xlsx", index=False, encoding="utf-8")

    # Get attachments
    def get_attachments(self, o_id, o_store_id, folder_path, pattern):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        attachments = message.Attachments

        if len(attachments) > 0:
            os.makedirs(folder_path, exist_ok=True)
            for attachment in attachments:
                if (
                    pattern.replace("*", "").lower() in str(attachment).lower()
                    or pattern == "*"
                ):
                    try:
                        attachment.SaveAsFile(
                            f"{folder_path}\\{str(attachment)}"
                        )

                    except Exception as ex:
                        print(ex.args)
                        try:
                            os.remove(f"{folder_path}\\{str(attachment)}")
                        except:
                            pass

    # Send new email
    def send_email(
        self, o_from, o_to, o_cc, o_subj, o_body, o_html_body, o_att_path
    ):
        email = self.outlook.CreateItem(0)
        email.To = o_to

        if o_cc:
            email.CC = o_cc
        email.Subject = o_subj

        if o_body:
            email.Body = o_body
        else:
            email.HTMLBody = o_html_body

        if o_from:
            email.SentOnBehalfOfName = o_from

        # Add attachments if any
        if o_att_path:
            for path in o_att_path:
                email.Attachments.Add(path)

        email.Send()

    # Reply to email
    def reply_to_email(
        self, o_id, o_store_id, o_body, o_html_body, o_att_path
    ):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        reply = message.Reply()
        if o_body:
            reply.Body = o_body
        else:
            reply.HTMLBody = o_html_body
        # Add attachments if any
        if o_att_path:
            for path in o_att_path:
                reply.Attachments.Add(path)

        reply.Send()

    # Save email as file
    def save_email(self, o_id, folder_path):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        filename = re.sub("\W", message.Subject)
        try:
            message.SaveAs(f"{folder_path}\\{filename}.msg")
        except Exception as ex:
            print(ex.args)
            pass

    # Move email to folder
    def move_email(self, o_id, o_store_id, o_new_folder):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        new_folder = self.ns.Folders[self.o_mailbox].Folders[o_new_folder]
        message.Move(new_folder)

    # Mark email item as read
    def mark_email(self, o_id, o_store_id):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        message.UnRead = False

    # Delete email item
    def delete_email(self, o_id, o_store_id):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        message.Delete()


# MAIN function
if __name__ == "__main__":
    # Get user
    user = os.getlogin()
    # Get today's date
    today = datetime.now()
    # Yesterday date
    yesterday = (
        (today - timedelta(days=1))
        .strftime("X%m-X%d-%Y")
        .replace("X0", "")
        .replace("X", "")
    )
    today = today.strftime("X%m-X%d-%Y").replace("X0", "").replace("X", "")

    # Set arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("--mailbox", help="Mailbox", default=None)
    parser.add_argument(
        "--mailbox_folder", help="Mailbox folder", default="Inbox"
    )
    parser.add_argument(
        "--mailbox_new_folder", help="Mailbox new folder", default="Inbox"
    )
    parser.add_argument(
        "--mail_filter",
        help="Mail filter",
        default=f"[SentOn] > '{yesterday} 12:00 AM' AND [SentOn] < '{today} 11:59 PM' AND [Unread] = True",
    )
    parser.add_argument(
        "--folder_path",
        help="Folder path",
        default=f"C:\\Users\\{user}\\Downloads",
    )
    parser.add_argument("--email_action", help="Email action", default=None)
    parser.add_argument("--email_id", help="Email ID", default=None)
    parser.add_argument(
        "--email_store_id", help="Email Store ID", default=None
    )
    parser.add_argument(
        "--att_pattern", help="Attachments pattern", default="*"
    )
    parser.add_argument("--from_address", help="Email Address", default=None)
    parser.add_argument("--to_address", help="Email Address", default=None)
    parser.add_argument("--cc_address", help="Email Address", default=None)
    parser.add_argument(
        "--email_subject", help="Email Subject", default="Subject empty"
    )
    parser.add_argument(
        "--email_body", help="Email Message", default="Body empty"
    )
    parser.add_argument(
        "--email_html_body",
        help="Email HTML Message",
        default="HTML body empty",
    )
    parser.add_argument("--att_path", help="Attachment path", default=None)

    # Get argument values
    args = parser.parse_args()
    o_mailbox = args.mailbox
    o_folder = args.mailbox_folder
    o_new_folder = args.mailbox_new_folder
    o_filter = args.mail_filter
    folder_path = args.folder_path
    action = args.email_action
    o_id = args.email_id
    o_store_id = args.email_store_id
    pattern = args.att_pattern
    o_from = args.from_address
    o_to = args.to_address
    o_cc = args.cc_address
    o_subj = args.email_subject
    o_body = args.email_body
    o_html_body = args.email_html_body
    o_att_path = args.att_path

    if o_att_path:
        if "," in o_att_path:
            o_att_path = o_att_path.split(",")
        else:
            o_att_path = [o_att_path]

    # Start script
    attempts = 1
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Delete error txt if exists
    if os.path.exists(f"{folder_path}\\error.txt"):
        os.remove(f"{folder_path}\\error.txt")

    # Attempt to run script only 3 times
    while attempts < 4:
        # Init object
        outlook = Outlook(o_mailbox, o_folder)
        try:
            # Choose action
            if action == "get_emails":
                outlook.get_emails(o_filter, folder_path)
            elif action == "get_attachments":
                outlook.get_attachments(o_id, o_store_id, folder_path, pattern)
            elif action == "send_email":
                outlook.send_email(
                    o_from, o_to, o_cc, o_subj, o_body, o_html_body, o_att_path
                )
            elif action == "reply_to_email":
                outlook.reply_to_email(
                    o_id, o_store_id, o_body, o_html_body, o_att_path
                )
            elif action == "save_email":
                outlook.save_email(o_id, o_store_id, folder_path)
            elif action == "mark_email":
                outlook.mark_email(o_id, o_store_id)
            elif action == "move_email":
                outlook.move_email(o_id, o_store_id, o_new_folder)

        # Catch exception
        except Exception as ex:
            print(f"Exception: {ex}\n\n")
            with open(f"{folder_path}\\error.txt", "w", encoding="utf-8") as f:
                f.write(str(ex))
                attempts += 1
                pass

        # Action if completed successfully
        else:
            with open(
                f"{folder_path}\\get_mail.txt", "w", encoding="utf-8"
            ) as f:
                f.write("Done")
                print("Done!")
                break

        # Close Outlook instance
        finally:
            outlook.close()
