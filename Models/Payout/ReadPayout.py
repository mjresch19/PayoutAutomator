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
        self.product_vendor = product_vendor.strip("'").replace("''", "'").title()
        self.product_type = product_type
        self.net_quantity = int(net_quantity)
        self.gross_sales = round(float(gross_sales),2)
        self.discounts = round(float(discounts),2)
        self.returns = round(float(returns),2)
        self.net_sales = round(float(net_sales),2)
        self.taxes = round(float(taxes),2)
        self.total_sales = round(float(total_sales),2)
        self.total_cost = round(float(total_cost),2)

    def get_distribution_type(self):
        
        if "(Original)" in self.product_vendor or "(original)" in self.product_vendor:
            return "Original"
        
        if "(Collab)" in self.product_vendor or "(collab)" in self.product_vendor:
            if self.product_type.count("(Collab)") + self.product_type.count("(collab)") > 1:
                return f"Group Collab ({self.product_type.count('Collab') + self.product_type.count('collab') + 1})"
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
