import json
import logging
import os
import pandas as pd
import sys
import win32com.client as win32

def main(thread, input, output, config: json, dir):
    # Configure logging
    now = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    logging.basicConfig(
        filename=os.path.join(output, f'{now}.log'),
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # open the outlook app, excel sheet, and html template
    try:
        # dispatch outlook instance
        outlook = win32.Dispatch('outlook.application')

        # get input data from excel
        input_data = pd.read_excel(input)
        
        # fill all empty values
        input_data = input_data.fillna('')
            
        # open html templates
        with open(f'{dir}\\src\\templates\\{config['template_file']}', 'r', encoding="utf8") as file:
            template=file.read()

            
    except Exception as e:
        logging.error(f'Error while initializing  script: {str(e)}')
        return 1
        
    logging.info(f'Initialized')

    return send_emails(thread, outlook, input_data, template, config)


# Group the data by the person's name
# for each person create table HTML and insert into template
def send_emails(thread, outlook, input_data, template, config: json):
  
    for name, group in input_data.groupby('Project Approver'):
        # initialize vars for email
        recipient = group.iloc[0]['Email Address ']
        logging.info(f'Sending new email to: {recipient}, {name}')
        subject = config['subject'].format(name=name)
        html = template
        
        # Create a list of dictionaries where each dictionary represents a row in the table
        table_html = ''

        try:
            for index, row in group.iterrows():
                table_html += '<tr>'
                for var, info in config['variables'].items():
                    value = row[info['column']]
                    if info['type'] == 'date':
                        value = pd.to_datetime(value).date()
                    elif info['type'] == 'float':
                        value = float(value)
                    elif info['type'] == 'int':
                        value = int(value)
                    elif info['type'] == 'str':
                        value = str(value)
                    value = info['format'].format(value)
                    table_html += f'<td nowrap="nowrap" style="width:144px;height:51px;"><p align="center" style="position: relative;">{value}</p></td>'                    
                table_html += '</tr>'
        except Exception as e:
            logging.error(f'Error creating email: {e}')
            logging.error('Please check that your excel file matches the template you selected')
            return 2

        try:
            # Replace the placeholder in the HTML template with the actual table
            html = html.replace('{rows}', table_html)
            
            # create the email in outlook and populate data
            mail = outlook.CreateItem(0)
            mail.To = recipient
            mail.Subject = subject
            mail.HtmlBody = html.format(name=name)
            # mail.Display(True)
            mail.Send()
        except Exception as e:
            logging.error(f'Error sending email to {recipient}: {str(e)}')
            continue

    return 0