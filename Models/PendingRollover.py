class PendingRollover:

    def __init__(
                    self,
                    product_title: str,  
                    product_vendor: str,
                    dist_type: str,
                    product_type: str, 
                    net_quantity: int, 
                    gross_sales: float, 
                    discounts: float, 
                    returns: float, 
                    net_sales: float, 
                    taxes: float, 
                    total_sales: float,
                    total_cost: float, 
                    processing_fee: float,
                    gross_profit: float,
                    origin: str,
                    ready: str,
                ):
        self.product_title = product_title.strip("'").replace("''", "'")
        self.product_vendor = product_vendor.strip("'").replace("''", "'").title()
        self.dist_type = dist_type
        self.product_type = product_type
        self.net_quantity = int(net_quantity)
        self.gross_sales = round(float(gross_sales),2)
        self.discounts = round(float(discounts),2)
        self.returns = round(float(returns),2)
        self.net_sales = round(float(net_sales),2)
        self.taxes = round(float(taxes),2)
        self.total_sales = round(float(total_sales),2)
        self.total_cost = round(float(total_cost),2)
        self.processing_fee = round(float(processing_fee),2)
        self.gross_profit = round(float(gross_profit),2)
        self.origin = origin
        self.ready = ready


def parse_pending_rollovers(pending_rollover_info, ynm_financial_info, yne_financial_info):
    '''
    This function parses the pending rollover information and separates it into the respective financial information lists.

    :param pending_rollover_info: list of lists containing the pending rollover information
    :param ynm_financial_info: list of lists containing the financial information for YNM
    :param yne_financial_info: list of lists containing the financial information for YNE

    :return: ynm_financial_info, yne_financial_info, carry_pending_rollovers
    '''
    
    carry_pending_rollovers = []

    for pending_rollover in pending_rollover_info:

        curr_pr = PendingRollover(
            pending_rollover[0],
            pending_rollover[1],
            pending_rollover[2], 
            pending_rollover[3], 
            pending_rollover[4], 
            pending_rollover[5],  
            pending_rollover[6],  
            pending_rollover[7], 
            pending_rollover[8],  
            pending_rollover[9],
            pending_rollover[10], 
            pending_rollover[11],  
            pending_rollover[12],
            pending_rollover[13],
            pending_rollover[14],
            pending_rollover[15]  
        )

        if curr_pr.ready.lower() == "y":

            if curr_pr.origin.upper() == "YNM":

                ynm_financial_info.append([curr_pr.product_title, curr_pr.product_vendor, curr_pr.dist_type, curr_pr.product_type, curr_pr.net_quantity, curr_pr.gross_sales,
                                            curr_pr.discounts, curr_pr.returns, curr_pr.net_sales, curr_pr.taxes,
                                            curr_pr.total_sales, curr_pr.total_cost, curr_pr.processing_fee,
                                            curr_pr.gross_profit])

            elif curr_pr.origin.upper() == "YNE":

                yne_financial_info.append([curr_pr.product_title, curr_pr.product_vendor, curr_pr.dist_type, curr_pr.product_type, curr_pr.net_quantity, curr_pr.gross_sales,
                                            curr_pr.discounts, curr_pr.returns, curr_pr.net_sales, curr_pr.taxes,
                                            curr_pr.total_sales, curr_pr.total_cost, curr_pr.processing_fee,
                                            curr_pr.gross_profit])

        else:
                
            carry_pending_rollovers.append(curr_pr)


    return ynm_financial_info, yne_financial_info, carry_pending_rollovers

