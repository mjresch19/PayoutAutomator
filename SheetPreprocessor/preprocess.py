import csv
import sys
import os
curr_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(curr_dir)

from Models.Payout.ReadPayout import ReadPayout
from Models.Payout.WritePayout import WritePayout

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
        next(reader)

        for row in reader:
            
            file_contents.append(row)


    return file_contents

def main():

    to_do_cost_list = []

    #Open raw report files
    ynm_file_path = '/YNM/PayoutAutomator/SheetPreprocessor/YNM_Sales.csv'
    ynm_financial_info = read_csv(ynm_file_path)

    yne_file_path = '/YNM/PayoutAutomator/SheetPreprocessor/YNE_Sales.csv'
    yne_financial_info = read_csv(yne_file_path)

    #Iterate throuh the raw YNM Information
    with open('/YNM/PayoutAutomator/SheetPreprocessor/YNM_Sales_Final.csv', 'w', newline='') as ynmcsvfile:
        ynmwriter = csv.writer(ynmcsvfile, delimiter=',')

        #Write header
        ynmwriter.writerow(["product_title", "product_name", "dist_type","product_type", "net_quantity", "gross_sales", "discounts", "returns", "net_sales", "taxes", "total_sales", "process_fee", "total_cost", "gross_profit"])

        for product in ynm_financial_info:

            read_payout = ReadPayout(
                product_title = product[0], 
                product_vendor = product[1], 
                product_type = product[2], 
                net_quantity = product[3], 
                gross_sales = product[4], 
                discounts = product[5], 
                returns = product[6], 
                net_sales = product[7], 
                taxes = product[8], 
                total_sales = product[9], 
                total_cost = product[10]
            )
            
            dist_type = read_payout.get_distribution_type()

            processing_fee = read_payout.calculate_processing_fee()

            gross_profit = read_payout.calculate_gross_profit()

            write_payout = WritePayout(
                product_title = read_payout.product_title,
                product_vendor = read_payout.product_vendor,
                distribution_type = dist_type,
                product_type = read_payout.product_type,
                net_quantity = read_payout.net_quantity,
                gross_sales = read_payout.gross_sales,
                discounts = read_payout.discounts,
                returns = read_payout.returns,
                net_sales = read_payout.net_sales,
                taxes = read_payout.taxes,
                total_sales = read_payout.total_sales,
                processing_fee = processing_fee,
                total_cost = read_payout.total_cost,
                gross_profit = gross_profit
            )

            ynmwriter.writerow([
                write_payout.product_title, 
                write_payout.product_vendor, 
                write_payout.distribution_type, 
                write_payout.product_type, 
                write_payout.net_quantity, 
                write_payout.gross_sales, 
                write_payout.discounts, 
                write_payout.returns, 
                write_payout.net_sales, 
                write_payout.taxes, 
                write_payout.total_sales, 
                write_payout.processing_fee,
                write_payout.total_cost, 
                write_payout.gross_profit
            ])


            if write_payout.total_cost == 0:
                
                to_do_cost_list.append([
                    "YNM", 
                    write_payout.product_title, 
                    write_payout.product_vendor, 
                    write_payout.distribution_type, 
                    write_payout.product_type, 
                    write_payout.net_quantity, 
                    write_payout.gross_sales, 
                    write_payout.discounts, 
                    write_payout.returns, 
                    write_payout.net_sales, 
                    write_payout.taxes, 
                    write_payout.total_sales, 
                    write_payout.processing_fee,
                    write_payout.total_cost, 
                    write_payout.gross_profit
                ])

    with open('/YNM/PayoutAutomator/SheetPreprocessor/YNE_Sales_Final.csv', 'w', newline='') as ynecsvfile:
        ynewriter = csv.writer(ynecsvfile, delimiter=',')

        #Write header
        ynewriter.writerow(["product_title", "product_name", "dist_type","product_type", "net_quantity", "gross_sales", "discounts", "returns", "net_sales", "taxes", "total_sales", "process_fee","total_cost", "gross_profit"])

        for product in yne_financial_info:

            read_payout = ReadPayout(
                    product_title = product[0], 
                    product_vendor = product[1], 
                    product_type = product[2], 
                    net_quantity = product[3], 
                    gross_sales = product[4], 
                    discounts = product[5], 
                    returns = product[6], 
                    net_sales = product[7], 
                    taxes = product[8], 
                    total_sales = product[9], 
                    total_cost = product[10]
                )
                
            dist_type = read_payout.get_distribution_type()

            processing_fee = read_payout.calculate_processing_fee()

            gross_profit = read_payout.calculate_gross_profit()

            write_payout = WritePayout(
                product_title = read_payout.product_title,
                product_vendor = read_payout.product_vendor,
                distribution_type = dist_type,
                product_type = read_payout.product_type,
                net_quantity = read_payout.net_quantity,
                gross_sales = read_payout.gross_sales,
                discounts = read_payout.discounts,
                returns = read_payout.returns,
                net_sales = read_payout.net_sales,
                taxes = read_payout.taxes,
                total_sales = read_payout.total_sales,
                processing_fee = processing_fee,
                total_cost = read_payout.total_cost,
                gross_profit = gross_profit
            )

            ynewriter.writerow([
                write_payout.product_title, 
                write_payout.product_vendor, 
                write_payout.distribution_type, 
                write_payout.product_type, 
                write_payout.net_quantity, 
                write_payout.gross_sales, 
                write_payout.discounts, 
                write_payout.returns, 
                write_payout.net_sales, 
                write_payout.taxes, 
                write_payout.total_sales, 
                write_payout.processing_fee,
                write_payout.total_cost, 
                write_payout.gross_profit
            ])


            if write_payout.total_cost == 0:
                
                to_do_cost_list.append([
                    "YNE", 
                    write_payout.product_title, 
                    write_payout.product_vendor, 
                    write_payout.distribution_type, 
                    write_payout.product_type, 
                    write_payout.net_quantity, 
                    write_payout.gross_sales, 
                    write_payout.discounts, 
                    write_payout.returns, 
                    write_payout.net_sales, 
                    write_payout.taxes, 
                    write_payout.total_sales, 
                    write_payout.processing_fee,
                    write_payout.total_cost, 
                    write_payout.gross_profit
                ])


    #Write to a third sheet that will be composed of all items that have no cost associated with them
    with open('/YNM/PayoutAutomator/SheetPreprocessor/No_Cost_Items.csv', 'w', newline='') as nocostitems:
        nocostitemswriter = csv.writer(nocostitems, delimiter=',')

        #Write Header
        nocostitemswriter.writerow(["Origin", "product_title", "product_name", "dist_type","product_type", "net_quantity", "gross_sales", "discounts", "returns", "net_sales", "taxes", "total_sales", "process_fee","total_cost", "gross_profit"])

        #Iterate through each no cost item and add to the sheet
        for item in to_do_cost_list:
            nocostitemswriter.writerow(item)

                      
main()