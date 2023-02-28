import unittest
import os
import re
import time
import win32com.client as wc

# Importing the code to be tested
from email_actions import Outlook


class TestOutlook(unittest.TestCase):
    # SetUp, executes before every function to test
    def setUp(self) -> None:
        self.user = os.getlogin()
        self.temp_path = f"C:/Users/{self.user}/Downloads"
        self.wait = 10
        self.mailbox = None
        self.mailbox_folder = None
        self.new_folder = "Desarrollos"
        # self.mailbox_name = "REMITOS CONFORMADOS"
        self.o = Outlook(self.mailbox, self.mailbox_folder)

    # TearDown, executes after every function to test
    def tearDown(self) -> None:
        outbox = self.o.ns.GetDefaultFolder(4)
        # Delete all items in the Outbox
        for item in outbox.Items:
            item.Delete()
        self.o.close

    # Test the __init__() method
    def test_init(self):
        self.assertIsInstance(self.o.outlook, wc.CDispatch)
        self.assertIsInstance(self.o.ns, wc.CDispatch)
        self.assertIsInstance(self.o.inbox, wc.CDispatch)

    # Test the get_emails() method
    def test_get_emails(self):
        self.o.get_emails("[Unread]=True", self.temp_path)
        self.assertTrue(os.path.exists(f"{self.temp_path}/df.xlsx"))
        os.remove(f"{self.temp_path}/df.xlsx")

    # Test the get_attachments() method
    def test_get_attachments_empty(self):
        self.o.inbox.Items.Add("IPM.Note")
        message = self.o.inbox.Items.GetLast()

        # Test the method when no attachments are present in the email
        self.o.get_attachments(
            message.EntryID, message.Parent.StoreID, ".", ".txt"
        )
        self.assertFalse(os.path.exists(f"{self.temp_path}/test.txt"))

    def test_get_attachments(self):
        self.o.inbox.Items.Add("IPM.Note")
        message = self.o.inbox.Items.GetLast()

        # create a temporary attachment file
        temp_file = os.path.join(self.temp_path, "test.txt")
        with open(temp_file, "w") as f:
            f.write("test file")

        # Add the attachment to the email
        message.Attachments.Add(temp_file)

        # Call the get_attachments() method and check that the attachment was saved
        self.o.get_attachments(
            message.EntryID, message.Parent.StoreID, ".", ".txt"
        )
        self.assertTrue(temp_file)

        # Remove the test file
        os.remove(temp_file)

    # Test send email() method
    def test_send_email(self):
        # Create a temporary file for attachment
        file_path = f"{self.temp_path}/test.txt"
        with open(file_path, "w") as f:
            f.write("This is a test file.")

        # Initialize EmailActions object and send email
        self.o.send_email(
            o_from="almolina@tenaris.com",
            o_to="almolina@tenaris.com",
            o_cc=None,
            o_subj="Test email",
            o_body="This is a test email.",
            o_html_body=None,
            o_att_path=[file_path],
        )

        # Remove the temporary file
        os.remove(file_path)
        time.sleep(self.wait)

        # Assert that the email was sent successfully
        outbox = self.o.ns.GetDefaultFolder(5)
        items = outbox.Items
        sent_email = None
        for item in items:
            if item.Subject == "Test email":
                sent_email = item
                break

        self.assertTrue(sent_email.Sent)
        time.sleep(self.wait)

    # Test reply_to_email() method
    def test_reply_to_email(self):
        time.sleep(self.wait)
        # Create a temporary file for attachment
        file_path1 = f"{self.temp_path}/test1.txt"
        with open(file_path1, "w") as f:
            f.write("This is a 2nd test file.")

        # Send email
        subject = "Test email"
        for message in self.o.inbox.Items:
            if message.Subject == subject:
                # Reply to the email
                reply_body = None
                reply_html_body = "<html><body><p>Thank you for your email.</p></body></html>"
                reply_attachments = [file_path1]
                self.o.reply_to_email(
                    message.EntryID,
                    message.Parent.StoreID,
                    reply_body,
                    reply_html_body,
                    reply_attachments,
                )
                break

        # Remove the temporary file
        os.remove(file_path1)
        time.sleep(self.wait)

        # Assert that the email was sent successfully
        outbox = self.o.ns.GetDefaultFolder(5)
        items = outbox.Items
        sent_email = None
        for item in items:
            if item.Subject == "RE: Test email":
                sent_email = item
                break

        self.assertTrue(sent_email.Sent)
        time.sleep(self.wait)

    # Test save_mail() method
    def test_save_mail(self):
        time.sleep(self.wait)

        subject = "RE: Test email"
        for message in self.o.inbox.Items:
            if message.Subject == subject:
                self.o.save_email(
                    message.EntryID, message.Parent.StoreID, self.temp_path
                )
                break

        # Check that the file was created
        subj = re.sub("\W", "_", subject)
        filename = os.path.join(self.temp_path, f"{subj}.msg")
        time.sleep(self.wait)
        self.assertTrue(os.path.isfile(filename))

        # Remove the temporary file
        os.remove(filename)

    # Test mark_email() method
    def test_mark_email(self):
        subject = "RE: Test email"
        email = None
        for message in self.o.inbox.Items:
            if message.Subject == subject:
                email = message
                self.o.mark_email(message.EntryID, message.Parent.StoreID)
                break

        time.sleep(self.wait)
        self.assertFalse(email.UnRead)

    # Test move_email() method
    def test_move_email(self):
        subject = "RE: Test email"
        email = None
        for message in self.o.inbox.Items:
            if message.Subject == subject:
                email = message
                self.o.move_email(
                    message.EntryID, message.Parent.StoreID, self.new_folder
                )
                break

        mailbox = self.mailbox
        new_folder = self.new_folder
        time.sleep(self.wait)
        self.o.close
        # assert that the email has been moved to the new folder
        self.o = Outlook(mailbox, new_folder)
        self.o.get_emails("[Unread]=True", self.temp_path)
        found = False
        for message in self.o.inbox.Items:
            if (
                email.Subject == message.Subject
            ):  # EntryID changes based on the folder where it's stored
                found = True
                break
        self.assertTrue(found)

    # Test delete_email() method
    def test_delete_email(self):
        subject = "Test email"
        email = None
        for message in self.o.inbox.Items:
            if message.Subject == subject:
                email = message
                self.o.delete_email(message.EntryID, message.Parent.StoreID)
                break

        # Try to get the deleted email, it should raise an error
        with self.assertRaises(Exception):
            self.o.ns.GetItemFromID(email.EntryID)


if __name__ == "__main__":
    # unittest.main()
    # Create a test suite
    test_suite = unittest.TestSuite()

    # Add the test cases to the test suite in the order you want them to be executed
    test_suite.addTest(TestOutlook("test_init"))
    test_suite.addTest(TestOutlook("test_get_emails"))
    test_suite.addTest(TestOutlook("test_get_attachments_empty"))
    test_suite.addTest(TestOutlook("test_get_attachments"))
    test_suite.addTest(TestOutlook("test_send_email"))
    test_suite.addTest(TestOutlook("test_reply_to_email"))
    test_suite.addTest(TestOutlook("test_save_mail"))
    test_suite.addTest(TestOutlook("test_mark_email"))
    test_suite.addTest(TestOutlook("test_move_email"))
    test_suite.addTest(TestOutlook("test_delete_email"))

    # Run the test suite
    unittest.TextTestRunner().run(test_suite)
