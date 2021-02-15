import email
import imaplib
import os
import pyperclip
import glob
import docx2txt
import time
from threading import Timer
from docx import Document
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# Host login information (Gmail)
host = 'imap.gmail.com'
username = '@gmail.com'
password = 'password'
attachment_dir = 'C:\\Users\\ahmed\\Desktop\\HutbeAttachments'

def run_timer():

# Download mail attachment to folder on PC

    def get_attachments(msg):
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()

            if bool(fileName):
                filePath = os.path.join(attachment_dir, fileName)
                with open(filePath, 'wb') as f:
                    f.write(part.get_payload(decode=True))

    my_inbox = []

# while loop if inbox is empty console prints searching for mail
# Only look for unread emails
    while len(my_inbox) == 0:
        print('searching for mail')

        def get_inbox():
            mail = imaplib.IMAP4_SSL(host)
            mail.login(username, password)
            mail.select("inbox")
            _, search_data = mail.search(None, 'UNSEEN')
            my_message = []
            for num in search_data[0].split():
                print(num)
                email_data = {}
                _, data = mail.fetch(num, '(RFC822)')
                # print(data)
                _, b = data[0]
                email_message = email.message_from_bytes(b)
                # print(email_message)
                for header in ['subject', 'to', 'from', 'date']:
                    print("{}: {}".format(header, email_message[header]))
                    email_data[header] = email_message[header]
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True)
                        email_data['body'] = body.decode()
                    elif part.get_content_type() == "text/html":
                        html_body = part.get_payload(decode=True)
                        email_data['html_body'] = html_body.decode()
                my_message.append(email_data)
                # pyperclip.copy(email_data['html_body'])
                raw = email.message_from_bytes(data[0][1])
                get_attachments(raw)
            return my_message

        if __name__ == '__main__':
            my_inbox = get_inbox()
            print(my_inbox)
        if len(my_inbox) > 0:
            break
        time.sleep(10)

    # Print Latest File
    # Search a folder in my directory called HutbeAttachments

    while len(os.listdir('C:\\Users\\ahmed\\Desktop\\HutbeAttachments')) == 0:
        files = os.listdir('C:\\Users\\ahmed\\Desktop\\HutbeAttachments')
        print("directory is empty")
        if len(os.listdir('C:\\Users\\ahmed\\Desktop\\HutbeAttachments')) > 0:
            print(files)
            break
        time.sleep(10)

    list_of_files = glob.glob(
        'C:\\Users\\ahmed\\Desktop\\HutbeAttachments\\*.docx')  # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    print(latest_file)

    document = Document(latest_file)
    for para in document.paragraphs:
        print(para.text)

    MY_TEXT = docx2txt.process(latest_file)
    # print(MY_TEXT)

    pyperclip.copy(MY_TEXT)

    os.remove(latest_file)

    # Automated Wordpress Login
    PATH = "C:\\Users\\ahmed\\Desktop\\chrome_driver_selenium\\chromedriver.exe"
    driver = webdriver.Chrome(PATH)

    USER = 'username'
    PASS = 'password'

    driver.get("WordPress Website URL")

    UsernameURL = driver.find_element_by_id('user_login')
    UsernameURL.send_keys(USER)

    PasswordURL = driver.find_element_by_id('user_pass')
    PasswordURL.send_keys(PASS)

    ButtonURL = driver.find_element_by_id('wp-submit')
    ButtonURL.click()

    # Automated Wordpress NewPost
    driver.get("WordPress Website URL/wp-admin/post-new.php")

    SwitchEditor = driver.find_element_by_class_name('wpb_switch-to-composer')
    SwitchEditor.click()

    # Automated Wordpress Paste
    elem = driver.find_element_by_name("content")
    actions = ActionChains(driver)
    actions.move_to_element(elem)
    actions.click(elem)
    elem.send_keys(Keys.CONTROL, 'v')

    # Automated Wordpress Post
    CheckBox = driver.find_element_by_id('in-category-5')
    CheckBox.click()

    def get_subject():
        mail = imaplib.IMAP4_SSL(host)
        mail.login(username, password)
        mail.select("inbox")
        _, search_data = mail.search(None, 'ALL')
        my_message = []
        for num in search_data[0].split():
            # print(num)
            email_data = {}
            _, data = mail.fetch(num, '(RFC822)')
            # print(data)
            _, b = data[0]
            email_message = email.message_from_bytes(b)
            # print(email_message)
            for header in ['subject']:
                print("{}".format(email_message[header]))
                pyperclip.copy(email_message['subject'])
            for part in email_message.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True)
                    email_data['body'] = body.decode()
                my_message.append(email_data)
            mail.store(num, '+FLAGS', '\\Deleted')
            mail.expunge()
            mail.close()
            mail.logout()
            return my_message

    if __name__ == '__main__':
        my_subject = get_subject()
        print(my_subject)

    EnterTitle = driver.find_element_by_name("post_title")
    actions = ActionChains(driver)
    actions.move_to_element(EnterTitle)
    actions.click(EnterTitle)
    EnterTitle.send_keys(Keys.CONTROL, 'v')

    PublishPost = driver.find_element_by_id('publish')
    PublishPost.click()

    Timer(25, run_timer).start()


run_timer()
