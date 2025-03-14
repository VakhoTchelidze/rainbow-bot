import configparser
import pandas as pd


def com_parser():
    # Initialize ConfigParser
    config = configparser.ConfigParser()

    # Read the ini file
    ini_file = 'C:\\Program Files (x86)\\USR-VCOM\\program.ini'  # Replace with your file path
    config.read(ini_file)

    # List to store extracted data
    data = []

    # Loop through all sections and collect router info
    for section in config.sections():
        if section.isdigit():  # Only process numbered sections
            router_info = {
                'Section': section,
                'Remarks': config.get(section, 'Remarks', fallback=''),
                'COMName': config.get(section, 'COMName', fallback=''),
                'NetProtocol': config.get(section, 'NetProtocol', fallback=''),
                'RemoteIP': config.get(section, 'RemoteIP', fallback=''),
                'RemotePort': config.get(section, 'RemotePort', fallback=''),
                'LocalPort': config.get(section, 'LocalPort', fallback=''),
                'RegID': config.get(section, 'RegID', fallback=''),
                'CloudId': config.get(section, 'CloudId', fallback=''),
                'CloudPw': config.get(section, 'CloudPw', fallback='')
            }
            data.append(router_info)

    # Convert to DataFrame and save to CSV
    df = pd.DataFrame(data)
    # output_csv = 'router_info.csv'
    # df.to_csv(output_csv, index=False)
    return df
