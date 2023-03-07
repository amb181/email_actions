"""
@Author: 60082363
Script to interact with Outlook app via MAPI using win32com.
TENARIS SOLUTIONS 2023.
"""
import logging
import argparse
import win32com.client
import os
import re
import pandas as pd
from datetime import datetime, timedelta


class Outlook:
    def __init__(self, o_mailbox, o_folder) -> None:
        # Creating an object for the outlook application.
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        self.o_mailbox = o_mailbox
        self.o_folder = o_folder
        logging.debug(f"Namespace: {self.ns} {datetime.now()}")
        logging.debug(f"Mailbox: {self.o_mailbox} {datetime.now()}")
        logging.debug(f"Mailbox folder: {self.o_folder} {datetime.now()}")
        # Update mailbox
        try:
            self.ns.SendAndReceive(True)
            self.outlook.Session.SendAndReceive(True)
        except Exception as ex:
            logging.error(f"ns.SendAndReceive init{ex.args} {datetime.now()}")

        if self.o_mailbox:
            self.inbox = self.ns.Folders[self.o_mailbox].Folders["Inbox"]
        else:
            self.inbox = self.ns.GetDefaultFolder(6)

        if self.o_folder != "Inbox" and self.o_folder is not None:
            for folder in self.inbox.Folders:
                if folder.Name == self.o_folder:
                    self.inbox = self.inbox.Folders[self.o_folder]
                    break
            print("Folder name not found.")
            logging.debug(f"Folder name not found. {datetime.now()}")
        logging.debug(f"__init__ - self.inbox: {self.inbox} {datetime.now()}")

    def close(self):
        # Sync all folders in the inbox and outbox
        try:
            self.ns.SendAndReceive(True)
            self.outlook.Session.SendAndReceive(True)
        except Exception as ex:
            logging.error(
                f"ns.SendAndReceive close {ex.args} {datetime.now()}"
            )

        # close the MAPI object
        self.outlook.Application.Quit()

        logging.debug(f"close - self.inbox: {self.inbox} {datetime.now()}")

    def clean_string(self, string):
        string = str(
            string.replace("\n", " ").replace("\t", " ").replace("\r", " ")
        )

        return string

    # End of auxiliar functions

    # Get email items
    def get_emails(self, o_filter, folder_path):
        self.o_filter = o_filter
        logging.debug(
            f"get_emails - self.o_filter: {self.o_filter} {datetime.now()}"
        )
        # Getting folder email items
        self.messages = self.inbox.Items
        filteredEmails = self.messages.Restrict(self.o_filter)
        # Creating an object to access items inside the inbox of outlook.
        self.messages = filteredEmails
        logging.debug(
            f"Items in folder:\n{[x.Subject for x in self.messages]} {datetime.now()}"
        )

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
            receiver = message.To
            cc = message.CC
            subject = message.Subject
            body = self.clean_string(message.Body)
            html_body = None
            try:
                if message.MeetingStatus == 1:
                    html_body = self.clean_string(message.HTMLBody)
            except Exception as ex:
                logging.error(
                    f"Message does not have MeetingStatus\n{ex.args} {datetime.now()}"
                )

            attachments_raw = message.Attachments
            attachments = [att.FileName for att in attachments_raw]
            attachments_count = len(attachments)
            received = message.ReceivedTime.strftime("%m/%d/%y %H:%M:%S")
            sent = message.SentOn.strftime("%m/%d/%y %H:%M:%S")
            sender = message.SenderName
            sender_add = message.SenderEmailAddress
            # Start format email
            if message.SenderEmailType == "EX":
                try:
                    sender_add = (
                        message.Sender.GetExchangeUser().PrimarySmtpAddress
                    )
                except Exception as ex:
                    logging.error(
                        f"Could not get Exchange usern{ex.args} {datetime.now()}"
                    )

            # End format sender email
            unread = message.UnRead

            new_row = [
                id,
                store_id,
                receiver,
                cc,
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
        df.to_excel(f"{folder_path}/df.xlsx", index=False, encoding="utf-8")
        logging.debug(f"get_emails - COMPLETED {datetime.now()}")

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
                            f"{folder_path}/{str(attachment)}"
                        )

                    except Exception as ex:
                        logging.error(
                            f"Could not download item {str(attachment)}\n{ex.args} {datetime.now()}"
                        )
                        os.remove(f"{folder_path}/{str(attachment)}")

        logging.debug(f"get_attachments - COMPLETED {datetime.now()}")

    # Send new email
    def send_email(
        self, o_from, o_to, o_cc, o_subj, o_body, o_html_body, o_att_path
    ):
        try:
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
                logging.debug(
                    f"send_email - o_from: {o_from} {datetime.now()}"
                )

            # Add attachments if any
            if o_att_path:
                logging.debug(
                    f"send_email - o_att_path: {o_att_path} {datetime.now()}"
                )
                for path in o_att_path:
                    email.Attachments.Add(path)

            email.Send()

            logging.debug(f"send_email - COMPLETED {datetime.now()}")

        except Exception as ex:
            logging.error(f"{ex.args} {datetime.now()}")

    # Reply to email
    def reply_to_email(
        self, o_id, o_store_id, o_body, o_html_body, o_att_path
    ):
        try:
            message = self.ns.GetItemFromID(o_id, o_store_id)
            reply = message.Reply()
            importance = "Low"
            if message.Importance == 1:
                importance = "Normal"
            elif message.Importance == 2:
                importance = "High"

            traceback = f"<html><body><br><b>From</b>: {message.Sender}\n<br><b>Sent</b>: {message.SentOn}\n<br><b>To</b>: {message.To}\n<br><b>Cc</b>: {message.CC}\n<br><b>Subject</b>: {message.Subject}\n<br><b>Importance</b>: {importance}\n\n</body></html>"
            if o_body:
                traceback = (
                    traceback.replace("<b>", "")
                    .replace("</b>", "")
                    .replace("<html><body>", "")
                    .replace("</body></html>", "")
                )
                reply.Body = f"{o_body}\n\n{traceback}\n{message.Body}"
            else:
                reply.HTMLBody = (
                    f"{o_html_body}\n\n{traceback}\n{message.HTMLBody}"
                )
            # Add attachments if any
            if o_att_path:
                for path in o_att_path:
                    reply.Attachments.Add(path)

            reply.Send()

            logging.debug(f"reply_to_email - COMPLETED {datetime.now()}")

        except Exception as ex:
            logging.error(f"{ex.args} {datetime.now()}")

    # Save email as file
    def save_email(self, o_id, o_store_id, folder_path):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        logging.debug(
            f"save_email - folder_path: {folder_path} {datetime.now()}"
        )
        filename = re.sub("\W", "_", message.Subject)
        try:
            message.SaveAs(os.path.join(folder_path, f"{filename}.msg"))
        except Exception as ex:
            logging.error(f"Could not save email\n{ex.args} {datetime.now()}")

        logging.debug(f"save_email - COMPLETED {datetime.now()}")

    # Move email to folder
    def move_email(self, o_id, o_store_id, o_new_folder):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        logging.debug(
            f"move_email - o_new_folder: {o_new_folder} {datetime.now()}"
        )

        try:
            if self.o_mailbox:
                message.Move(
                    self.ns.Folders[self.o_mailbox]
                    .Folders["Inbox"]
                    .Folders[o_new_folder]
                )
            else:
                message.Move(self.ns.GetDefaultFolder(6).Folders[o_new_folder])
            logging.debug(f"move_email - COMPLETED {datetime.now()}")

        except Exception as ex:
            logging.error(f"Could not move email\n{ex.args} {datetime.now()}")

    # Mark email item as read
    def mark_email(self, o_id, o_store_id):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        try:
            message.UnRead = False
            logging.debug(f"mark_email - COMPLETED {datetime.now()}")
        except Exception as ex:
            logging.error(f"Could not mark email\n{ex.args} {datetime.now()}")

    # Delete email item
    def delete_email(self, o_id, o_store_id):
        message = self.ns.GetItemFromID(o_id, o_store_id)
        try:
            message.Delete()
            logging.debug(f"delete_email - COMPLETED {datetime.now()}")
        except Exception as ex:
            logging.error(
                f"Could not delete email\n{ex.args} {datetime.now()}"
            )


# MAIN function
if __name__ == "__main__":
    # Get user
    user = os.getlogin()
    logging.basicConfig(
        filename=f"C:/Users/{user}/AppData/Local/Temp/email_actions.log",
        level=logging.DEBUG,
    )
    # Get today's date
    today = datetime.now()
    # Yesterday date
    yesterday = (
        (today - timedelta(days=3))
        .strftime("%Y-X%m-X%d")
        .replace("X0", "")
        .replace("X", "")
    )
    today = today.strftime("%Y-X%m-X%d").replace("X0", "").replace("X", "")

    # Set arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("--mailbox", help="Mailbox", default=None)
    parser.add_argument(
        "--mailbox_folder", help="Mailbox folder", default=None
    )
    parser.add_argument(
        "--mailbox_new_folder", help="Mailbox new folder", default=None
    )
    parser.add_argument(
        "--mail_filter",
        help="Mail filter",
        default=f"[SentOn] > '{yesterday} 12:00 AM' AND [SentOn] < '{today} 11:59 PM' AND [Unread] = True",
    )
    parser.add_argument(
        "--folder_path",
        help="Folder path",
        default=f"C:/Users/{user}/Downloads",
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
    parser.add_argument("--email_subject", help="Email Subject", default=None)
    parser.add_argument("--email_body", help="Email Message", default=None)
    parser.add_argument(
        "--email_html_body",
        help="Email HTML Message",
        default=None,
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
    if os.path.exists(f"{folder_path}/error.txt"):
        os.remove(f"{folder_path}/error.txt")

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
            elif action == "delete_email":
                outlook.delete_email(o_id, o_store_id)

        # Catch exception
        except Exception as ex:
            logging.error(f"Could perform action\n{ex.args} {datetime.now()}")
            with open(f"{folder_path}/error.txt", "w", encoding="utf-8") as f:
                logging.error(f"{ex.args} {datetime.now()}")
                f.write(str(ex))
                attempts += 1

        # Action if completed successfully
        else:
            with open(
                f"{folder_path}/get_mail.txt", "w", encoding="utf-8"
            ) as f:
                f.write("Done")
                print("Done!")
                break

        # Close Outlook instance
        finally:
            outlook.close()

        logging.debug(f"PROCESS LOOP FINISHED {datetime.now()}\n")
    logging.debug(f"PROCESS FINISHED {datetime.now()}\n\n\n")
