def create_outlook_email(to_address: str, subject: str, body: str):
    """
    Create and display a new Outlook email via COM (Windows Outlook).
    """
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError("Outlook/COM interface unavailable (not on Windows or pywin32 not installed).") from e
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_address or ""
    mail.Subject = subject or ""
    mail.Body = body or ""
    mail.Display()
