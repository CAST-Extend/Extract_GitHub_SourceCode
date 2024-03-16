from argparse import ArgumentParser
import logging
import os
import subprocess
import openpyxl
from datetime import datetime

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
        print(f"Error reading Excel file: {e}\n")
        logging.error(f"Error reading Excel file: {e}\n")
    return data

def create_directory_if_not_exists(directory_path):
    if not os.path.exists(directory_path):
        try:
            os.makedirs(directory_path)
            print(f"Directory '{directory_path}' created successfully.\n")
            logging.info(f"Directory '{directory_path}' created successfully.\n")
            return False
        except OSError as e:
            print(f"Error: {e}\n")
            logging.error(f"Error: {e}\n")
    else:
        print(f"Directory '{directory_path}' already exists.\n")
        logging.info(f"Directory '{directory_path}' already exists.\n")
        return True

def download_and_save_code(application_name, repository_url, server_location, access_token, repository_owner):

    application_name_directory = server_location+'\\'+application_name
    application_name_directory_flag = create_directory_if_not_exists(application_name_directory)

    repository_name = repository_url.split('/')[-1]
    repository_path = application_name_directory+'\\'+repository_name
    repository_path_flag = create_directory_if_not_exists(repository_path)
    if repository_path_flag:
        print(f"Repository '{repository_name}' already downloaded to '{repository_path}'.\n")
        logging.info(f"Repository '{repository_name}' already downloaded to '{repository_path}'.\n")
    else:
        try:
            clone_url = f"https://{access_token}@github.com/{repository_owner}/{repository_name}.git"
            # Running the git clone command
            subprocess.run(['git', 'clone', clone_url, repository_path], check=True)

            print(f"Repository '{repository_name}' downloaded successfully to '{repository_path}'.\n")
            logging.info(f"Repository '{repository_name}' downloaded successfully to '{repository_path}'.\n")
        except subprocess.CalledProcessError as e:
            print(f"An error occurred: {e}")
            logging.error(f"An error occurred: {e}")

def download_in_batches(repositories, batch_number, access_token, repository_owner):
    for repository in repositories:
        if repository[2] == batch_number:
            download_and_save_code(repository[0], repository[1], repository[3], access_token, repository_owner)

if __name__ == "__main__":

    parser = ArgumentParser()
 
    parser.add_argument('-excel_file', '--excel_file', required=True, help='Excel File Name')
    parser.add_argument('-batch','--batch', required=True, help='Batch Number')
    parser.add_argument('-access_token', '--access_token', required=True, help='Access Token')
    parser.add_argument('-repository_owner', '--repository_owner', required=True, help='Repository Owner')

    args=parser.parse_args()

    os.makedirs('GitHub_clone_logs', exist_ok=True)
    datetime_now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    # Set up logging
    log_file = os.path.join('GitHub_clone_logs',f"GitHub_clone_logs_{datetime_now}.log")
    logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


    data = read_excel_data(args.excel_file)
    # print(data)

    download_in_batches(data, int(args.batch), args.access_token, args.repository_owner)
