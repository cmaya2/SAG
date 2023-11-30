import csv
import xmltodict
from datetime import *


class Convert_940:
    def __init__(self, CSV):
        self.CSV = CSV

    def parse_csv(self):

        with open(self.CSV, 'r', errors="ignore", encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            orders = {}

            for row in reader:
                order_number = row['ClientOrderNumber']
                customer_name = row['CustomerCode']
                purchase_order_date = row['OrderDate'].replace('/', '-')
                requested_dilvery_date = row['StartDate'].replace('/', '-')
                cancel_date = row['CancelDate'].replace('/', '-')
                routing = row['ReserveCode'].strip()
                carrier_scac = row['CarrierSCAC']
                empty_value = ''
                order_status = 'New'

                if order_number not in orders:
                    if carrier_scac == 'FDEN':
                        carrier_scac = 'FDEG'

                    orders[order_number] = {
                        'OrderHeader': {
                            'Facility': row.get('FacilitySite'),
                            'Client': row.get('ClientCode'),
                            'DepositorOrderNumber': row['ClientOrderNumber'],
                            'OrderStatus': order_status,
                            'PurchaseOrderNumber': row.get('CustomerPo'),
                            'ShipTo': {
                                'Name': row.get('ShipToName'),
                                'Code': row.get('ShipToName'),
                                'Address1': row.get('ShipToAddress1'),
                                'Address2': row.get('ShipToAddress2'),
                                'City': row.get('ShipToCity'),
                                'State': row.get('ShipToState'),
                                'ZipCode': row.get('ShipToPostalCode'),
                                'Country': row.get('ShipToCountryCode'),
                                'ContactPhone': row.get('ShipToPhone')
                            },
                            'Dates': {
                                'PurchaseOrderDate': purchase_order_date,
                                'RequestedDeliveryDate': requested_dilvery_date,
                                'CancelDate': cancel_date
                            },
                            'ReferenceInformation': {
                                'CustomerName': row.get('CustomerCode'),
                                'VendorNumber': empty_value,
                                'AccountNumber': row.get('BillToAccountNum')
                            },
                            'ShippingInstructions': {
                                'ShipmentMethodOfPayment': empty_value,
                                'TransportationMethod': empty_value,
                                'CarrierCode': carrier_scac,
                                'Routing': routing
                            },
                            'Messages': empty_value
                        },
                        'OrderDetail': {
                            'OrderLine': []
                        }
                    }

                order_line = {
                    'OrderLineNumber': row.get('LineNumber'),
                    'ItemNumber': row.get('ItemCode'),
                    'CaseUPC': empty_value,
                    'BuyerItemNumber': empty_value,
                    'OrderedQuantity': row.get('OrderQuantity'),
                    'QuantityUnitOfMeasure': row.get('UOM'),
                    'ItemDescription': empty_value,
                    'ReferenceNumber': empty_value
                }
                orders[order_number]['OrderDetail']['OrderLine'].append(order_line)

                # Convert each order dictionary to XML using xmltodict and write to separate XML files
                xml_data = xmltodict.unparse({'Order': orders[order_number]}, short_empty_elements=True,
                                             pretty=True)

                with open("C:\\FTP\\GPAEDIProduction\\Integral\\In\\940_79_" + order_number + "_" +
                          customer_name + "_" + datetime.now().strftime("%Y%m%d%H%M%S" + ".xml"), "w") \
                        as acknowledgement_file:
                    acknowledgement_file.write(xml_data)
                with open("C:\\FTP\\GPAEDIProduction\\TG1-SAG\\Out\\Archive\\940\\940_79_" + order_number + "_" +
                          customer_name + "_" + datetime.now().strftime("%Y%m%d%H%M%S" + ".xml"), "w") \
                        as acknowledgement_file:
                    acknowledgement_file.write(xml_data)

