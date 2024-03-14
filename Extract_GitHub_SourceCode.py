from argparse import ArgumentParser
import os
from git import Repo
import openpyxl

def read_excel_data(file_path):
    data = []
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            server_location = row[3].replace('\\\\', '\\')
            data.append((row[0], row[1], row[2], server_location))
        workbook.close()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
    return data

def create_directory_if_not_exists(directory_path):
    if not os.path.exists(directory_path):
        try:
            os.makedirs(directory_path)
            print(f"Directory '{directory_path}' created successfully.\n")
        except OSError as e:
            print(f"Error: {e}")
    else:
        print(f"Directory '{directory_path}' already exists.\n")

def download_and_save_code(application_name, repository_url, server_location):

    application_name_directory = server_location+'\\'+application_name
    create_directory_if_not_exists(application_name_directory)

    repository_name = repository_url.split('/')[-1]
    repository_path = application_name_directory+'\\'+repository_name
    create_directory_if_not_exists(repository_path)
    
    # Clone the repository
    Repo.clone_from(repository_url, repository_path)
    print(f"Repository '{repository_name}' downloaded successfully to '{repository_path}'.\n")

def download_in_batches(repositories, batch_number):
    for repository in repositories:
        if repository[2] == batch_number:
            download_and_save_code(repository[0], repository[1], repository[3])

if __name__ == "__main__":

    parser = ArgumentParser()
 
    parser.add_argument('-excel_file', '--excel_file', required=True, help='Excel File Name')
    parser.add_argument('-batch','--batch', required=True, help='Batch Number')

    args=parser.parse_args()

    data = read_excel_data(args.excel_file)
    # print(data)

    download_in_batches(data, int(args.batch))
