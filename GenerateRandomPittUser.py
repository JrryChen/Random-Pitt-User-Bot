from __future__ import print_function  # Python 2/3 compatibility
import random
import win32com.client as win32com  # For Outlook


def lookUpUser(username):
    outlook = win32com.Dispatch("Outlook.Application").GetNamespace("MAPI")
    gal = outlook.Session.GetGlobalAddressList()
    entries = gal.AddressEntries
    ae = entries[username]
    email_address = None

    if 'EX' == ae.Type:
        eu = ae.GetExchangeUser()
        email_address = eu.PrimarySmtpAddress

    if 'SMTP' == ae.Type:
        email_address = ae.Address

    print('Email address: ', email_address)


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
    return (username)


def main():
    pittUser = generateRandomPittUser()
    lookUpUser(pittUser)


if __name__ == "__main__":
    main()
