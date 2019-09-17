####################################
# Author: Jon Škerlj               #
# Company: Mirada Medical Ltd      #
# Date (last modified): 17/09/2019 #
####################################


import requests
import sys
import getpass

from functions import *
from Credentials import credentials


def main():
    #uncomment for other users
    #Input details user needs to insert
    username = input("Enter your username (e.g. john.smith): ")
    password = getpass.getpass("Enter your password: ")

    credentials = username + ":" + password + "@"

    absPath = input("Enter your project directory: ")

    project = input("Enter project name (e.g. NMX_4_6): ")

    # input for url of the server
    # if the whole url is: http://10.0.0.99:8080/mirada_clouds/FUXD_4_0/F-SSRS-1
    # url user has to input equals to 10.0.0.99:8080/mirada_clouds
    url = input("Enter project url (what is between http:// and project name): ")

    resp = requests.get('http://{}{}/rest/1/{}/cat'.format(credentials, url, project))

    while(resp.status_code!=200):
        print("\nYour input data was incorrect, please try again!\n")

        username = input("Enter your username (e.g. john.smith): ")
        password = getpass.getpass("Enter your password: ")

        credentials = username + ":" + password + "@"

        absPath = input("Enter your project directory: ")

        project = input("Enter project name (e.g. NMX_4_6): ")

        url = input("Enter project url (what is between http:// and project name): ")

        resp = requests.get('http://{}{}/rest/1/{}/cat'.format(credentials, url, project))

    print("\nYour input data is correct!\n")

    # Commented variables for testing features (avoiding user input)
    # comment user input (above) and uncomment varibales below
    # variables probably need some modifications - absPath & credentials
    # absPath = "C:\\Users\\jon.skerlj\\Desktop\\TTM test XD"
    # credentials = "jon.skerlj:*****@"
    # project = "FUXD_4_0"
    # url = "10.0.0.99:8080/mirada_clouds"1


    # folders in the mentioned directory
    Folders = structureSet(absPath + "\\set.xml")

    # Creating a list of objects - each stores name and path of the folder
    MMclass = classList(Folders, absPath)

    # List that is going to store levels of each folder in order of user input
    level = []
    idx = 0
    # List that is going to store what elements are to be skipped if user says so
    skip=[]
    # Getting level of test/requirment for all folders
    skipBool = False
    for ele in MMclass:
        # Getting new directory (inside one of the 'mm' folders)
        directory = os.listdir(ele.path)

        # Reading the name of the main folder
        name = readTitle(ele.path + "\\" + "document.xml")
        # Getting folder type - either Test or Requirement
        typ = readType(ele.path + "\\" + "document.xml")

        if (typ == "Requirement"):
            print(name)
            lvl = ""

            # Loop that makes sure that one of the two answers is chosen
            while (lvl != "SSRS" and lvl != "SURS"):
                lvl = input("Please enter level of requirement (SSRS/SURS):")

                if(lvl != "SSRS" and lvl != "SURS"):
                    print("Your input was incorrect, try again!")


            level.append(lvl)
        else:
            print(name)
            lvl = ""
            # Loop that makes sure that one of the two answers is chosen
            while (lvl != "SSTS" and lvl != "SUTS"):
                lvl = input("Please enter level of test (SSTS/SUTS):")

                if (lvl != "SSTS" and lvl != "SUTS"):
                    print("Your input was incorrect, try again!")

            level.append(lvl)

        #Check if the folder already exists
        #Only checks main folders (right under SSTS/SUTS/SSRS/SURS)
        resp = requests.get(
            'http://{}{}/rest/1/{}/item/F-{}-1'.format(credentials, url, project, lvl),
            data={'[children]': 'yes'})
        # Checking if that folder already exists
        for abc in resp.json()["itemList"]:
            if (abc["title"] == name):
                print("Folder with the same name already exists!")

                yn = ""
                while (yn != "Y" and yn != "N"):
                    yn = input("Do you want to continue? (Y/N):")
                if (yn == "Y"):
                    yn2 = ""
                    while (yn2 != "Y" and yn2 != "N"):
                        yn2 = input("Do you want to skip uploading this folder? (Y/N):")
                    if(yn2 == "Y"):
                        # Getting the index of an element that needs to be skipped
                        skip.append(idx)
                        skipBool = True
                    else:
                        print("Creating another instance of the folder.")
                    break
                else:
                    sys.exit(0)
        idx += 1

    #Getting the id of contents element of a folder in Matrix
    folderID = getMatrixID("FOLDER", credentials, project, url)

    # Checks if user said wanted to skip some folders that already exist
    if(skipBool):
        # idx = 0
        temp = MMclass
        MMclass = []
        i=0
        for idx, ele in enumerate(temp):
            print(idx)
            if(skip[i]==idx):
                i+=1
            else:
                MMclass.append(ele)


    # Looping through each main folder
    for ele in MMclass:
        # Getting new directory (inside one of the 'mm' folder)
        directory = os.listdir(ele.path)

        # Reading the name of the main folder
        name = readTitle(ele.path + "\\" + "document.xml")

        # Getting folder prefix - old naming systen
        prefix = redPrefix(ele.path + "\\" + "document.xml")

        # allocating corresponding id with right level - needed for inserting file data onto the website
        # desID - description id number of a field on a website where you display information
        # level[count] - "SSRS/SURS/SSTS/SUTS" - user allocated into the level list
        desID = getMatrixID(level[count], credentials, project, url)

        data = {
            "reason":"test",
            # Name of the folder
            "label": name,
            # Name of the folder where we want to create a new folder
            "parent": "F-{}-1".format(level[count])
        }

        # Post request that creates a new folder on the website
        resp = requests.post("http://{}{}/rest/1/{}/folder".format(credentials, url, project), data=data)

        # Getting the identification of the new created folder - used for creating new subfolders in a created folder
        index = resp.json()["serial"]

        # Getting folders in the subdirectory
        Folders2 = structureFolder(ele.path + "\\structure.xml")
        print(Folders2)

        # Checks if there are any folders left
        if not Folders2:
            # if no folders found ends the current stage of loop - moves to the next element
            # Nothing gets uploaded to the main folder contents element

            # Going to the next element in a loop
            continue
        else:
            # More folders found - create a list of objects again
            Folders2Class = classList(Folders2, ele.path)
            # Using recursive function to go through directory
            directorySearch(Folders2Class, absPath, index, level[count], desID, prefix, project, credentials, folderID, url)


        #Counting the number of loops - at the end increment by one
        count = count + 1

if __name__ == '__main__':
    main()