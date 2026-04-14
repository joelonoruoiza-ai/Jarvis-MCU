import win32com.client

class OutlookMailAccess:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox

    def get_accounts(self):
        return [account.DisplayName for account in self.namespace.Accounts]

    def get_unread_count(self):
        return len(self.inbox.Items.Restrict("[Read] = false"))

    def get_recent_messages(self, count=10):
        recent_items = self.inbox.Items
        recent_items.Sort("[ReceivedTime]", True)
        return [recent_items[i] for i in range(min(count, recent_items.Count))]

    def get_unread_messages(self):
        return [item for item in self.inbox.Items if item.Unread]

    def search_mail(self, subject_or_sender):
        return [item for item in self.inbox.Items if subject_or_sender.lower() in item.Subject.lower() or subject_or_sender.lower() in item.SenderName.lower()]

    def read_message(self, message):
        return message.Body or message.HTMLBody
