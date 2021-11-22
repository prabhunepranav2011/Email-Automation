import glob
import os
import shutil
import win32com.client
import datetime

def saveattachemnts(messages,today,path): 

    for message in messages:
         if message.Unread and message.Senton.date() == today:
                attachments = message.Attachments

            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))

                if message.Unread:
                    message.Unread = False
            break 

path = r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles"  #Enter the desired Path, download locatation
today = datetime.date.today()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#Use GetNameSpace ("MAPI") to return the Outlook NameSpace object from the Application object.

inbox = outlook.GetDefaultFolder(6)
Filter = "[SenderEmailAddress] = 'prabhune.pranav2011@gmail.com'"     #Enter the sender Email Address

messages = inbox.Items.Restrict(Filter)      
saveattachemnts(messages,today,path)

directory1 = "Labeler 66733"     #Enter the 3 Labelers
directory2 = "Labeler 0002"
directory3 = "Labeler "
parent_dir = r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles"   #Location where you want to create folders

path1 = os.path.join(parent_dir, directory1)
path2 = os.path.join(parent_dir, directory2)
path3 = os.path.join(parent_dir, directory3)

os.mkdir(path1)
os.mkdir(path2)
os.mkdir(path3)

src_folder = r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles"     #SOurce Folder
dst_folder1 = r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles\Labeler 66733\\"     #destination foldernames
dst_folder2 = r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles\Labeler 0002\\"
#dst_folder3 = r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles\Labeller \\"

pattern1 = src_folder + r"\66733*"
pattern2 = src_folder + r"\0002*"
#pattern3 = src_folder + #enter pattern here

patterns = [pattern1, pattern2, pattern3] 
destinations = [dst_folder1, dst_folder2, dst_folder3]

for i in range(len(patterns)):
    #os.chdir(src_folder)
    for file in glob.iglob(patterns[i], recursive=False):
    # extract file name form file path
        file_name = os.path.basename(file)
        #print(file_name)
        shutil.move(file, destinations[i] + file_name)
        #print('Moved:', file)

