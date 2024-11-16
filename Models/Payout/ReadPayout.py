class ReadPayout:
    def __init__(
                    self,
                    product_title: str,  
                    product_vendor: str,
                    product_type: str, 
                    net_quantity: int, 
                    gross_sales: float, 
                    discounts: float, 
                    returns: float, 
                    net_sales: float, 
                    taxes: float, 
                    total_sales: float, 
                    total_cost: float, 
                ):
        
        self.product_title = product_title.strip("'").replace("''", "'") 
        self.product_vendor = product_vendor.strip("'").replace("''", "'")
        self.product_type = product_type
        self.net_quantity = net_quantity
        self.gross_sales = gross_sales
        self.discounts = discounts
        self.returns = returns
        self.net_sales = net_sales
        self.taxes = taxes
        self.total_sales = total_sales
        self.total_cost = total_cost


    def get_distribution_type(self):
        
        if "(Original)" in self.product_vendor or "(original)" in self.product_vendor:
            return "Original"
        
        if "(Collab)" in self.product_vendor or "(collab)" in self.product_vendor:
            if self.product_type.count("(Collab)") + self.product_type.count("(collab)") > 1:
                return "Group Collab"
            else:
               return "Collab"
            
        if "(Commercial)" in self.product_vendor or "(commercial)" in self.product_vendor:
            return "Commercial"

        if "(Charity)" in self.product_vendor or "(charity)" in self.product_vendor:
            return "Charity"
        
        return "Unknown"
    
    def calculate_processing_fee(self):

        return round(float(self.total_sales) * .03,2)
    

    def calculate_gross_profit(self):

        return self.total_sales - self.calculate_processing_fee() - self.total_cost
