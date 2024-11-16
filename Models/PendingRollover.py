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