class PendingRollover:

    def __init__(self, origin, artist, item, distribution_type, total_value, processing_fee, cost, profit, ready):
        self.origin = origin
        self.artist = artist
        self.item = item
        self.distribution_type = distribution_type
        self.total_value = total_value
        self.processing_fee = processing_fee
        self.cost = cost
        self.profit = profit
        self.ready = ready