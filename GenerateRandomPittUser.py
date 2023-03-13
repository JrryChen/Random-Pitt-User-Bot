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


def sendEmail(email):
    outlook = win32com.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = 'Email Bot Test'
    mail.Body = 'ECON 0150 - Email Bot Test'
    mail.Display()
    print('Email Displayed')
    # mail.Send()
    # print('Email Sent')


def main():
    # testPittUser = 'JZC23'
    pittUser = generateRandomPittUser()
    email = lookUpUser(pittUser)
    sendEmail(email)

    # emails = ['amm517@pitt.edu', 'jzc23@pitt.edu', 'jpl86@pitt.edu']  # Testing
    # for email in emails:
    #     print(email)
    #     sendEmail(email)


if __name__ == "__main__":
    main()
