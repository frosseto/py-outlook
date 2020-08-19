from win32com.client import Dispatch
import win32com.client


def novo_email(html, fanexo=None, assunto='', to_list='', cc_list='', bcc_list='', send=False):
    '''Create, insert and send html.email using outlook'''
    obj = win32com.client.Dispatch("Outlook.Application")
    mail = obj.CreateItem(0x0)

    mail.Subject = assunto
    if not (fanexo is None):
        mail.Attachments.Add(fanexo)
    mail.HtmlBody = html
    mail.To = to_list
    mail.CC = cc_list
    mail.BCC = bcc_list
    mail.display()
    if send: mail.Send()