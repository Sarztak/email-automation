import win32com.client as win32

def open_outlook():
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return namespace
    
def get_draft_folder(namespace):
    drafts_folder = namespace.GetDefaultFolder(16)
    return drafts_folder

def format_draft():
    ...

def clean_data():
    ...

def send_email():
    ...




def main():
    namespace = open_outlook()
    draft_folder = get_draft_folder(namespace)
    print(draft_folder)


if __name__ == "__main__":
    main()
