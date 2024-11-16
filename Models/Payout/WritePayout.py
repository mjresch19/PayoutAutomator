class WritePayout:
    def __init__(
                self, 
                product_title,  
                product_vendor,
                distribution_type,
                product_type, 
                net_quantity, 
                gross_sales, 
                discounts, 
                returns, 
                net_sales, 
                taxes, 
                total_sales, 
                process_fee, 
                total_cost, 
                gross_profit
                ):
        
        self.product_title = product_title
        self.product_vendor = product_vendor
        self.distribution_type = distribution_type
        self.product_type = product_type
        self.net_quantity = net_quantity
        self.gross_sales = gross_sales
        self.discounts = discounts
        self.returns = returns
        self.net_sales = net_sales
        self.taxes = taxes
        self.total_sales = total_sales
        self.process_fee = process_fee
        self.total_cost = total_cost
        self.gross_profit = gross_profit
