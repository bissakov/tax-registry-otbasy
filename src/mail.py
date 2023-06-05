import win32com.client

from data_structures import EmailInfo
from src.utils import kill_all_processes


class Email:
    def __init__(self) -> None:
        kill_all_processes(proc_name='OUTLOOK.EXE')
        self.email = EmailInfo()
        self.outlook = win32com.client.Dispatch('Outlook.Application')

    def send(self):
        mail = self.construct_mail()
        mail.Send()
        self.outlook.Quit()

    def construct_mail(self):
        mail = self.outlook.CreateItem(0)
        mail.To = self.email.recepient
        mail.Subject = self.email.subject
        mail.Body = self.email.body
        if self.email.attachment:
            mail.Attachments.Add(self.email.attachment)
        return mail


if __name__ == '__main__':
    Email().send()
