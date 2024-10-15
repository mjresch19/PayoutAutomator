# from ..utilities.database import database
import csv
import sys
sys.path.append('/YNM/YNMPayoutAutomator/utilities')
from database import database

def read_csv(file_path):
    '''
    Translate csv contents into a python datatype

    @param file_path: The path to the csv file
    @return A list of lists containing the contents of the csv file
    '''

    file_contents = []

    with open(file_path, 'r') as file:
        reader = csv.reader(file)

        #Skip the header
        header = next(reader)

        for row in reader:
            
            file_contents.append(row)


    return file_contents

def main():

    to_do_cost_list = []

    #Open raw report files
    ynm_file_path = '/YNM/YNMPayoutAutomator/SheetPreprocessor/YNM_Sept.csv'
    ynm_financial_info = read_csv(ynm_file_path)

    yne_file_path = '/YNM/YNMPayoutAutomator/SheetPreprocessor/YNE_Sept.csv'
    yne_financial_info = read_csv(yne_file_path)

    #Iterate throuh the raw YNM Information
    with open('/YNM/YNMPayoutAutomator/SheetPreprocessor/YNM_Sept_Final.csv', 'w', newline='') as ynmcsvfile:
        ynmwriter = csv.writer(ynmcsvfile, delimiter=',')

        #Write header
        ynmwriter.writerow(["product_title", "product_name", "dist_type","product_type", "net_quantity", "gross_sales", "discounts", "returns", "net_sales", "taxes", "total_sales", "process_fee", "total_cost", "gross_profit"])

        #Iterate through each product that was sold this month for YNM
        for product in ynm_financial_info:

            product_vendor  = product[1].strip("'").replace("''", "'")
            product_name = product[0].strip("'").replace("''", "'") #Clean up the string to be readable in json
            product_type = product[2]
            net_quantity = product[3]
            gross_sales = product[4]
            discounts = product[5]
            returns = product[6]
            net_sales = product[7]
            taxes = product[8]
            total_sales = product[9]
            process_fee = round(float(product[9]) * .03,2)
            total_cost = float(product[10])
            gross_profit = float(product[9]) - process_fee - total_cost

            if total_cost == 0:
                to_do_cost_list.append(["YNM", product_name, product_vendor, dist_type, product_type, net_quantity, gross_sales, discounts, returns, net_sales, taxes, total_sales, process_fee, total_cost, gross_profit])
            
            ynmwriter.writerow([product_name, product_vendor, dist_type,product_type, net_quantity, gross_sales, discounts, returns, net_sales, taxes, total_sales, process_fee, total_cost, gross_profit])

    with open('/YNM/YNMPayoutAutomator/SheetPreprocessor/YNE_Sept_Final.csv', 'w', newline='') as ynecsvfile:
        ynewriter = csv.writer(ynecsvfile, delimiter=',')

        #Write header
        ynewriter.writerow(["product_title", "product_name", "dist_type","product_type", "net_quantity", "gross_sales", "discounts", "returns", "net_sales", "taxes", "total_sales", "process_fee","total_cost", "gross_profit"])

        #Iterate through each product that was sold this month for YNM
        for product in yne_financial_info:

            product_vendor  = product[1].strip("'").replace("''", "'")
            product_name = product[0].strip("'").replace("''", "'") #Clean up the string to be readable in json
            product_type = product[2]
            net_quantity = product[3]
            gross_sales = product[4]
            discounts = product[5]
            returns = product[6]
            net_sales = product[7]
            taxes = product[8]
            total_sales = product[9]
            process_fee = round(float(product[9]) * .03,2)
            total_cost = float(product[10])
            gross_profit = float(product[9]) - process_fee - total_cost

            if total_cost == 0:
                to_do_cost_list.append(["YNE", product_name, product_vendor, dist_type, product_type, net_quantity, gross_sales, discounts, returns, net_sales, taxes, total_sales, process_fee, total_cost, gross_profit])
            
            ynewriter.writerow([product_name, product_vendor, dist_type,product_type, net_quantity, gross_sales, discounts, returns, net_sales, taxes, total_sales, process_fee, total_cost, gross_profit])

    #Write to a third sheet that will be composed of all items that have no cost associated with them
    with open('/YNM/YNMPayoutAutomator/SheetPreprocessor/No_Cost_Items.csv', 'w', newline='') as nocostitems:
        nocostitemswriter = csv.writer(nocostitems, delimiter=',')

        #Write Header
        nocostitemswriter.writerow(["Origin", "product_title", "product_name", "dist_type","product_type", "net_quantity", "gross_sales", "discounts", "returns", "net_sales", "taxes", "total_sales", "process_fee","total_cost", "gross_profit"])

        #Iterate through each no cost item and add to the sheet
        for item in to_do_cost_list:
            nocostitemswriter.writerow(item)

                      
main()