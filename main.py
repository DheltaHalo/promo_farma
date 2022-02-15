import os
import json
import yaml
import tkinter
import pandas as pd
from tkinter import filedialog
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.client import OAuth2Credentials

os.chdir(os.path.dirname(os.path.abspath(__file__)))    


def xlsx_to_csv(target_folder: str, files_location: str):
    """
    Gets all the files from target_folder that have .xlsx ending and
    saves them in a 'files' folder in the 'files_location' as a .csv file.

    Parameters
    ----------
    target_folder: str, folder where your .xlsx files are
    files_location: str, folder where the .csv files will be saved

    """
    local = os.walk(target_folder, topdown=True)

    # We also save all files path's
    path_list = []
    
    for (root,dirs,files) in local:
        for file in files:
            f_path = os.path.join(root, file) # Gets the file relative path
            if '.xlsx' in f_path:
                df = pd.read_excel(f_path) # Reads the file into a dataframe
                new_path = os.path.join(files_location, 'files') # New folder where files will be stored

                # Now we create a folder if it doesn't exist
                if not os.path.isdir(new_path): 
                    os.makedirs(new_path)
                
                new_path = os.path.join(new_path, os.path.basename(f_path)).replace('xlsx', 'csv')
                path_list.append(new_path)
                df.to_csv(new_path)
                
    return path_list


def upload_google_drive(files_paths: str, files_location):
    """
    Uploads files from files_path to google drive using the Google API and it's associated credentials 
    """
    # First we configure the place where the credentials will be stored
    creds_path = os.path.join(files_location, 'credentials')

    if not os.path.isdir(creds_path):
        os.makedirs(creds_path)

    with open('settings.yaml', 'r+') as file:
        yaml_list = yaml.safe_load(file)
        yaml_list['save_credentials_file'] = os.path.join(creds_path, 'creds.dat')

    with open('settings.yaml', 'w+') as file:
        yaml.dump(yaml_list, file)

    # Then we aunthenticate using those credentials
    gauth = GoogleAuth()
    drive = GoogleDrive(gauth)  

    file_list = drive.ListFile({'q': "'{}' in parents and trashed=false".format('1Ej57MaG3oF-eXHcxi8YtT5fpIxkJa-In')}).GetList()
    name_list = [x['title'] for x in file_list]
   
    # Now we upload the files from files_paths
    for file in files_paths:
        name = os.path.basename(file)

        if name not in name_list:
            gfile = drive.CreateFile({'parents': [{'id': '1Ej57MaG3oF-eXHcxi8YtT5fpIxkJa-In'}]})
            # Rename the file
            gfile['title'] = name
            # Read file and set it as the content of this instance.
            gfile.SetContentFile(file)
            gfile.Upload() # Upload the file. 

    # And we get the links
    file_list = drive.ListFile({'q': "'{}' in parents and trashed=false".format('1Ej57MaG3oF-eXHcxi8YtT5fpIxkJa-In')}).GetList()
    link_list = []
    for file in file_list:
        link_list.append(file['webContentLink'])

    with open('links.txt', 'w+') as file:
        for link in link_list:
            file.write(link + "\n")


def getPath():
    root = tkinter.Tk()
    root.withdraw()
    path = filedialog.askdirectory(initialdir=os.path.expanduser('~'))
    root.destroy()

    return path


def main():
    """
    First it converst the input .xlsx files to .csv files.
    Then the program uploads them too google drive.
    """    
    files_path = os.path.expanduser('~/PromoFarma')
    excel_files = getPath()

    if excel_files == ():
        excel_files = '.'

    if not os.path.isdir(files_path):
        os.makedirs(files_path)

    files = xlsx_to_csv(excel_files, files_path)  
    upload_google_drive(files, files_path)
    
if __name__ == '__main__':
    main()