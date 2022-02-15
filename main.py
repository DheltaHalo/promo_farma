import os
import tempfile
import pandas as pd
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def xlsx_to_csv(target_folder: str, tmp_location: str):
    """
    Gets all the files from target_folder that have .xlsx ending and
    saves them in a 'files' folder in the 'tmp_location' as a .csv file.

    Parameters
    ----------
    target_folder: str, folder where your .xlsx files are
    tmp_location: str, folder where the .csv files will be saved

    """
    local = os.walk(target_folder, topdown=True)
    safe_folder = tmp_location

    # We also save all files path's
    path_list = []
    
    for (root,dirs,files) in local:
        for file in files:
            f_path = os.path.join(root, file) # Gets the file relative path
            if '.xlsx' in f_path:
                df = pd.read_excel(f_path) # Reads the file into a dataframe
                new_path = os.path.join(safe_folder, 'promofarma_files') # New folder where files will be stored

                # Now we create a folder if it doesn't exist
                if not os.path.isdir(new_path): 
                    os.makedirs(new_path)
                
                new_path = os.path.join(new_path, os.path.basename(f_path))
                path_list.append(new_path)
                df.to_csv(new_path.replace('.xlsx', '.csv'))
                
    return path_list


def upload_google_drive():
    gauth = GoogleAuth()           
    drive = GoogleDrive(gauth)  
    file_list = drive.ListFile({'q': "'{}' in parents and trashed=false".format('1cIMiqUDUNldxO6Nl-KVuS9SV-cWi9WLi')}).GetList()
    for file in file_list:
        print('title: %s, id: %s' % (file['title'], file['id']))


def main():
    """
    First it converst the input .xlsx files to .csv files.
    Then the program uploads them too google drive.
    """    

    tempdir = tempfile.gettempdir()
    # files = xlsx_to_csv('.', tempdir)    
    upload_google_drive()
    
    input()

if __name__ == '__main__':
    main()