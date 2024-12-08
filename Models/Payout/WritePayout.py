# import sys
# import os
# curr_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# sys.path.append(curr_dir)

# # Add the directory containing PayoutAutomator to the path
# sys.path.append(os.path.join(curr_dir, 'PayoutAuto'))

# from PayoutAuto.datalookups import artist_lookup, item_lookup, name_extract

class WritePayout:
    def __init__(
                    self, 
                    product_title: str,  
                    product_vendor: str,
                    distribution_type: str,
                    product_type: str, 
                    net_quantity: int, 
                    gross_sales: float, 
                    discounts: float, 
                    returns: float, 
                    net_sales: float, 
                    taxes: float, 
                    total_sales: float, 
                    processing_fee: float, 
                    total_cost: float, 
                    gross_profit: float
                ):
        
        self.product_title = product_title
        self.product_vendor = product_vendor.title()
        self.distribution_type = distribution_type
        self.product_type = product_type
        self.net_quantity = net_quantity
        self.gross_sales = gross_sales
        self.discounts = discounts
        self.returns = returns
        self.net_sales = net_sales
        self.taxes = taxes
        self.total_sales = total_sales
        self.processing_fee = processing_fee
        self.total_cost = total_cost
        self.gross_profit = gross_profit

    def extract_artist(self):
        '''
        This function will extract the artist from the product_vendor attribute and isolate it from
        the distribution type that is listed on the original product_vendor attribute
        '''
        
        if "(" in self.product_vendor:

            self.product_vendor = self.product_vendor[:self.product_vendor.index("(") - 1]

    def identify_artist(self, artist_info, item_info):

        if self.product_vendor.title() in artist_info:

            self.product_vendor = self.product_vendor.title()

        else:

            #Search for aliases
            search_vendor = artist_lookup(self.product_vendor, artist_info)

            if search_vendor:

                self.product_vendor = search_vendor

            else:
                #Search for items that may have artist info
                search_vendor_item = item_lookup(self.product_title, item_info)

                if search_vendor_item:

                    self.product_vendor = search_vendor_item

                else:

                    print("YNM ARTIST NOT FOUND, TAKE ACTION. SKIPPING FOR NOW:",
                           self.product_vendor.title(), "==>",self.product_vendor, "(" + self.product_title + ")")

