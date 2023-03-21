from __future__ import print_function  # Python 2/3 compatibility
import random
import win32com.client as win32com  # For Outlook


def generateRandomPittUser():  # Generate a random Pitt user
    # Generate three random letters
    first = chr(random.randint(65, 90))
    second = chr(random.randint(65, 90))
    third = chr(random.randint(65, 90))

    # Generate 2 random numbers
    num1 = random.randint(0, 9)
    num2 = random.randint(0, 9)

    username = first + second + third + str(num1) + str(num2)
    # print(username)
    return username


def lookUpUser(username):
    print('Looking up user: ', username)
    outlook = win32com.Dispatch("Outlook.Application").GetNamespace("MAPI")
    gal = outlook.Session.GetGlobalAddressList()
    entries = gal.AddressEntries
    ae = entries[username]
    print('Address Entry: ', ae)  # dubugging
    email_address = None

    if 'EX' == ae.Type:
        print('UserType: Exchange')
        eu = ae.GetExchangeUser()
        if eu == None:
            print('User not found')
            return
        else:
            email_address = eu.PrimarySmtpAddress

    if 'SMTP' == ae.Type:
        print('UserType: SMTP')
        email_address = ae.Address

    print('Email Address: ', email_address)
    return email_address


# def sendEmail(email, control):
#     if (email == None):
#         print('No Email')
#         return
#     if control:
#         link = "https://forms.gle/gQ8R7eWjn8wNBqve8"
#     else:
#         link = "https://forms.gle/P7c9QFMRBqW3FN8V8"
#     outlook = win32com.Dispatch("Outlook.Application")
#     mail = outlook.CreateItem(0)
#     mail.To = email
#     mail.Subject = 'Brief Individual Decision-Making Survey Request'
#     mail.Body = """Hello,\n
# I hope this email finds you well. I represent a group of undergraduate students of the University. We are reaching out to you to ask for your participation in a brief survey about individual decision-making. It will only take less than a minute to complete and your participation will be greatly appreciated!
# The email address I am using to contact you was obtained through a random username generator, and I assure you that your privacy and anonymity will be maintained throughout the survey.
# Your participation in this survey will help us gain valuable insights into individual decision-making, and we would be grateful if you could spare a minute to complete it. The survey consists of a few multiple-choice questions and should not take more than a minute to finish.
# To participate in the survey, please click on the following link: """ + link + """\n
# If you have any questions or concerns, please do not hesitate to contact us. Thank you in advance for your participation!   
# Best regards,
# Jerry Chen, Amanda Magill, Jared Lawrence
#     """

#     # mail.Display()
#     # print('Email Displayed')
#     mail.Send()
#     print('Email Sent')
#     return

def sendEmail(email, control):
    if (email == None):
        print('No Email')
        return
    if control:
        link = "https://forms.gle/gQ8R7eWjn8wNBqve8"
    else:
        link = "https://forms.gle/P7c9QFMRBqW3FN8V8"
    
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = 'Brief Individual Decision-Making Survey Request'
    
    # HTML body with an image
    mail.HTMLBody = """\
    <html>
      <head></head>
      <body>
        <p>Hello,</p>
        <p>I hope this email finds you well. I represent a group of undergraduate students of the University. We are reaching out to you to ask for your participation in a brief survey about individual decision-making. It will only take less than a minute to complete and your participation will be greatly appreciated!</p>
        
        <p>The email address I am using to contact you was obtained through a random username generator, and I assure you that your privacy and anonymity will be maintained throughout the survey.</p>
        <p>Your participation in this survey will help us gain valuable insights into individual decision-making, and we would be grateful if you could spare a minute to complete it. The survey consists of a few multiple-choice questions and should not take more than a minute to finish.</p>
        <p>To participate in the survey, please click on the following link:</p>
        <p><a href="{0}">{0}</a></p>
        <p>If you have any questions or concerns, please do not hesitate to contact us. Thank you in advance for your participation!</p>
        <p>Best regards,<br>
        Jerry Chen, Amanda Magill, Jared Lawrence</p>
        <p><img src="https://example.com/image.png" alt="Survey Image"></p>
      </body>
    </html>
    """.format(link)

    mail.Display()
    print('Email Displayed')
    # mail.Send()
    # print('Email Sent')
    return

def main():
    # testPittUser = 'JZC23'
    # pittUser = generateRandomPittUser()
    # email = lookUpUser(testPittUser)
    # sendEmail('jzc23@pitt.edu', True)

    emails = ['amm517@pitt.edu', 'jzc23@pitt.edu', 'jpl86@pitt.edu']  # Testing
    for email in emails:
        print(email)
        sendEmail(email, True)


if __name__ == "__main__":
    main()
