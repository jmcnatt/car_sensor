import argparse
import boxsdk
import configparser
import datetime
import errno
import json
import os
import pandas

def parse_args():
    """Parse the command line args"""
    argparser = argparse.ArgumentParser()
    argparser.add_argument('-c', '--config', default='car_sensor.conf', help='Path to config file')
    argparser.add_argument('-o', '--output', default='cars.json', help='Output file location')
    return argparser.parse_args()

def parse_config(path):
    config = configparser.ConfigParser()
    try:
        if os.path.exists(path):
            config.read(path)
        else:
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), path)
    except:
        print(f"Could not open config file {path}")
        raise
    return config

def auth_box(config):
    """Authenticates to Box using a Server app with Client Crediental Grant"""
    try:
        client_id = config['box']['client_id']
        client_secret = config['box']['client_secret']
        enterprise_id = config['box']['enterprise_id']
    except:
        print(f'Could not parse config file section for Box')
        raise

    try:
        auth = boxsdk.CCGAuth(
            client_id = client_id,
            client_secret = client_secret,
            enterprise_id = enterprise_id,
        )
        client = boxsdk.Client(auth)
        print(f'Authenticated to Box as app user {client.user().get().id}')
        return client
    except:
        print(f'Failed to authenticate to Box')

def parse_excel_workbook(workbook):
    """Parses Excel workbook content in bytes"""
    try:
        sheet = 'Properties'

        # Open the Excel file as a data structure
        dataframe = pandas.read_excel(io=workbook, sheet_name=sheet)
        data = {}

        for index, row in dataframe.iterrows():
            data[row['Key']] = row['Value']

        return data
    except:
        print(f"Unable to parse Excel workbook")
        raise

def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""

    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat()
    raise TypeError ("Type %s not serializable" % type(obj))

def save_json(results, output):
    """Saves the json results to the specified output file"""
    with open(output, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=4, sort_keys=True, default=json_serial)

if __name__ == '__main__':
    # Parse args
    args = parse_args()

    # Parse config
    config = parse_config(args.config)

    # Auth to Box
    box = auth_box(config)

    # Create initial json
    results = {'last_data_refresh': datetime.datetime.now().isoformat()}

    # Loop through each file key and parse the Excel workbook
    for key in config['cars']:
        file_id = config['cars'][key]
        print(f'Processing {key} file {file_id}')
        file_content = box.file(file_id).content(file_version=None, byte_range=None)
        workbook_results = parse_excel_workbook(file_content)
        results[key] = workbook_results

    # Return the results as JSON
    save_json(results, args.output)
    print(f"Output saved to {args.output}")
