class PendingRollover:

    def __init__(
                    self, 
                    origin: str,
                    artist: str, 
                    item: str, 
                    distribution_type: str, 
                    total_value: float, 
                    processing_fee: float, 
                    cost: float, 
                    profit: float, 
                    ready: bool
                ):
        self.origin = origin
        self.artist = artist
        self.item = item
        self.distribution_type = distribution_type
        self.total_value = total_value
        self.processing_fee = processing_fee
        self.cost = cost
        self.profit = profit
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
            origin = pending_rollover[0],
            artist = pending_rollover[1],
            item = pending_rollover[2],
            distribution_type = pending_rollover[3],
            total_value = pending_rollover[4],
            processing_fee = pending_rollover[5],
            cost = pending_rollover[6],
            profit = pending_rollover[7],
            ready = pending_rollover[8],
        )

        if curr_pr.ready.lower() == "y":

            if curr_pr.origin.upper() == "YNM":

                ynm_financial_info.append([curr_pr.item, curr_pr.artist, curr_pr.distribution_type, '', '', '', '',
                                            '', '', '', curr_pr.total_value, curr_pr.processing_fee, curr_pr.cost,
                                            curr_pr.profit])

            elif curr_pr.origin.upper() == "YNE":

                yne_financial_info.append([curr_pr.item, curr_pr.artist, curr_pr.distribution_type, '', '', '', '',
                                            '', '', '', curr_pr.total_value, curr_pr.processing_fee, curr_pr.cost,
                                            curr_pr.profit])

        else:
                
            carry_pending_rollovers.append(curr_pr)


    return ynm_financial_info, yne_financial_info, carry_pending_rollovers

