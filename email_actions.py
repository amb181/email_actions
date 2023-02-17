import win32com.client
import os
import sys
import pandas as pd


class Outlook:
    def __init__(self, o_mailbox, o_folder, o_mark_read, o_filter) -> None:
        self.o_mailbox = o_mailbox
        self.o_folder = o_folder
        self.o_mark_read = o_mark_read
        self.o_filter = o_filter

        # Creating an object for the outlook application.
        self.outlook = win32com.client.Dispatch("Outlook.Application")

        self.ns = self.outlook.GetNamespace("MAPI")

        # Creating an object to access outlook folder
        if self.o_mailbox == " ":
            self.inbox = self.ns.GetDefaultFolder(6)
        else:
            self.inbox = self.ns.Folders[self.o_mailbox].Folders[self.o_folder]

        # Filter email
        raw_messages = self.inbox.Items
        filteredEmails = raw_messages.Restrict(self.o_filter)
        # Creating an object to access items inside the inbox of outlook.
        self.messages = filteredEmails

    def close(self):
        # close the MAPI object
        self.outlook.Application.Quit()

    def get_messages(self):
        # Columns for dataframe
        df_columns = [
            "id",
            "receiver",
            "cc",
            "subject",
            "body",
            "html_body",
            "attachments",
            "attachments_count",
            "received",
            "sent",
            "sender",
            "sender_add",
            "unread",
        ]

        # To iterate through inbox emails using inbox.Items object.
        df_rows = []
        for message in self.messages:
            id = message.EntryID
            receiver = message.To
            cc = message.CC
            subject = message.Subject
            body = message.Body.replace("\n", " ").replace("\t", " ")
            html_body = message.HTMLBody.replace("\n", " ").replace("\t", " ")
            attachments_raw = message.Attachments
            attachments = [att.FileName for att in attachments_raw]
            attachments_count = len(attachments)
            received = message.ReceivedTime.strftime("%m/%d/%y %H:%M:%S")
            sent = message.SentOn.strftime("%m/%d/%y %H:%M:%S")
            sender = message.SenderName
            # Format sender email
            sender_add = message.SenderEmailAddress
            try:
                sender_add = self.ns.CreateRecipient(sender_add)
                sender_add.Resolve()
                sender_address_entry = sender_add.AddressEntry
                exchange_user = sender_address_entry.GetExchangeUser()
                sender_add = exchange_user.PrimarySmtpAddress
            except:
                pass

            # End
            unread = message.UnRead

            new_row = [
                id,
                receiver,
                cc,
                subject,
                body,
                html_body,
                "|".join(attachments),
                attachments_count,
                received,
                sent,
                sender,
                sender_add,
                unread,
            ]
            df_rows.append(new_row)

            if self.o_mark_read == "True":
                message.UnRead = False

        df = pd.DataFrame(df_rows, columns=df_columns)
        df.to_excel(f"{folder_path}\\df.xlsx", index=False, encoding="utf-8")

    def get_attachments(self, id, folder_path):
        for message in self.messages:

            if message.EntryID == id:
                attachments = message.Attachments

                if len(attachments) > 0:
                    os.makedirs(folder_path, exist_ok=True)

                    for attachment in attachments:
                        try:
                            attachment.SaveAsFile(
                                f"{folder_path}\\{str(attachment)}"
                            )

                        except Exception as ex:
                            print(ex.args)
                            message.UnRead = True
                            pass

                        else:
                            if self.o_mark_read == "True":
                                message.UnRead = False

                break


if __name__ == "__main__":
    # print(sys.argv)
    o_mailbox = sys.argv[1]
    o_folder = sys.argv[2]
    o_filter = sys.argv[3].replace("'", "'")
    o_mark_read = sys.argv[4]
    folder_path = sys.argv[5]  # .replace("\\", "\\\\")
    action = sys.argv[6]
    o_id = sys.argv[7]

    attempts = 1

    while attempts < 4:
        try:
            # Init object
            outlook = Outlook(o_mailbox, o_folder, o_mark_read, o_filter)
            # Choose action
            if action == "get_items":
                outlook.get_messages()
            elif action == "get_attachments":
                outlook.get_attachments(
                    o_id,
                    folder_path,
                )

        except Exception as ex:
            with open(f"{folder_path}\\error.txt", "w", encoding="utf-8") as f:
                f.write(str(ex))
                attempts += 1
                pass

        else:
            with open(
                f"{folder_path}\\get_mail.txt", "w", encoding="utf-8"
            ) as f:
                f.write("Done")
                print("Done!")
                break

        finally:
            outlook.close()
