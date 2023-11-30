import xml.etree.ElementTree as et
from datetime import datetime
import csv
from send_944_email import *
from prettytable import PrettyTable


class Convert_944:

    def __init__(self, XML):
        self.XML = XML

    def parseXMLemail(self):
        date = str(datetime.now())
        date = date.replace('-', '').replace(':', '').replace(' ', '').split('.')
        t = PrettyTable(
            ['ReceiptNumber', 'ItemNumber', 'ItemDescription', 'LotNumber', 'PurchaseOrderNumber',
             'PackQuantity',
             'ExpectedQuantity',
             'ReceivedQuantity', 'TotalCases', 'Variances'])

        tree = et.parse(self.XML)
        root = tree.getroot()
        # Extract ReceiptNumber from ReceiptHeader
        receipt_number = root.find('.//ReceiptNumber').text if root.find('.//ReceiptNumber') is not None else ''
        depositor_order_number = root.find('.//DepositorOrderNumber').text if root.find(
            './/DepositorOrderNumber') is not None else ''

        # Open a CSV file for writing
        with open(rf'C:\FTP\GPAEDIProduction\TG1-SAG\Out\Archive\CSV\Receipt_79_{receipt_number}_{date[0]}.csv', 'w',
                  newline='', encoding='utf-8') as csvfile:
            # Define the CSV writer
            fieldnames = ['ReceiptNumber', 'ItemNumber', 'ItemDescription', 'LotNumber', 'PurchaseOrderNumber',
                          'PackQuantity',
                          'ExpectedQuantity', 'ReceivedQuantity', 'CSQuantity', 'Variance']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            # Write the CSV header
            writer.writeheader()

            # Iterate over each item in ReceiptDetail
            for item in root.findall('.//ReceiptDetail/Item'):
                item_number = item.findtext('ItemNumber', '')
                item_description = item.findtext('ItemDescription', '')
                lot_number = item.findtext('LotNumber', '')
                purchase_order_number = item.findtext('PurchaseOrderNumber', '')
                pack_quantity = item.findtext('PackQuantity', '')
                shipped_quantity = item.findtext('ShippedQuantity', '')
                received_quantity = item.findtext('ReceivedQuantity', '')
                # print(shipped_quantity)

                RQ = int(received_quantity)
                PQ = int(pack_quantity)
                variance = int(shipped_quantity) - int(received_quantity)
                CSQuantity = 'N/A' if (RQ / PQ) == (0) else round(RQ / PQ)
                t.add_row(
                    [receipt_number, item_number, item_description, lot_number, purchase_order_number, pack_quantity,
                     shipped_quantity, received_quantity, CSQuantity, variance])
                # Extract data from XML
                item_data = {
                    'ReceiptNumber': receipt_number,
                    'ItemNumber': item.findtext('ItemNumber', ''),
                    'ItemDescription': item.findtext('ItemDescription', ''),
                    'LotNumber': item.findtext('LotNumber', ''),
                    'PurchaseOrderNumber': item.findtext('PurchaseOrderNumber', ''),
                    'PackQuantity': item.findtext('PackQuantity', ''),
                    'ExpectedQuantity': item.findtext('ShippedQuantity', ''),
                    'ReceivedQuantity': item.findtext('ReceivedQuantity', ''),
                    'CSQuantity': CSQuantity,
                    'Variance': variance
                }

                writer.writerow(item_data)

        try:
            subject = f'Receiving Report: {receipt_number}'
            client = 'Syndicated Apparel Group'
            logo = 'https://www.gpalogisticsgroup.com/wp-content/uploads/2020/06/GPA-Logistics-Logo.jpg'
            message = f"""
                <html>
                <head>
                  <style>
                    table {{
                      font-family: Arial, sans-serif;
                      border-collapse: collapse;
                      width: 100%;
                    }}
                    th, td {{
                      border: 1px solid #dddddd;
                      text-align: left;
                      padding: 8px;
                    }}
                    th {{
                      background-color: #f2f2f2;
                    }}
                    .highlight {{
                        color: #FF9900;
                    }}
                  </style>
                </head>
                <body>
                    <p>Hello <strong>{client}</strong>,</p>
                    <p>Please find the attached Receiving Report.</p>
                    <br>
                  {t.get_html_string()}
                    <br>

                    <div style="font-family: Arial, sans-serif; font-size: 12px;">
                        <p><strong>Thank you,</strong></p>
                        <p><strong>GPA TEAM</strong></p>
                        <p><strong>GPA Logistics Group, Inc.</strong></p>
                        <p>1600 S. Baker Ave.<br>
                        Ontario, CA 91761</p>
                        <hr>
                        <p><a href="http://www.gpalogisticsgroup.com">www.gpalogisticsgroup.com</a></p>
                        <p><a href="http://www.gpaglobal.net">www.gpaglobal.net</a></p>
                            <img src={logo} alt="GPA Logistics Logo" style="width:150px;">
                        <p>Shenzhen | Hong Kong | New York | <span class ="highlight">Los Angeles</span> | San Francisco | Miami | Cambridge UK | Dublin | Mexico | Vietnam | Thailand | Malaysia</p>
                    </div>

                </body>
                </html>
                """
            recipients = [
                'avelazquez@gpalogisticsgroup.com',
                'gpaops27@gpalogisticsgroup.com',
                'gpacs16@gpalogisticsgroup.com',
                'Robert@ringoffireclothing.com',
                'Josefina.H@ringoffireclothing.com',
                'Inbound7379@gpalogisticsgroup.com'
            ]


            send_email_csv(
                rf'C:\FTP\GPAEDIProduction\TG1-SAG\Out\Archive\CSV\Receipt_79_{receipt_number}_{date[0]}.csv',
                subject, message, recipients)
        except Exception as e:
            print(e)