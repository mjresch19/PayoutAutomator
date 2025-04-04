import json

with open('/YNM/PayoutAutomator/Data/margins.json') as fp:
    anamoly_info = json.load(fp)

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
        gross_profit: float,
        gross_margin: float = 0.0
    ):
        self.product_title = product_title.strip("'").replace("''", "'")
        self.product_vendor = product_vendor.strip("'").replace("''", "'").title()
        self.product_type = product_type
        self.net_quantity = int(net_quantity)
        self.gross_sales = round(float(gross_sales), 2)
        self.discounts = round(float(discounts), 2)
        self.returns = round(float(returns), 2)
        self.net_sales = round(float(net_sales), 2)
        self.taxes = round(float(taxes), 2)
        self.total_sales = round(float(total_sales), 2)
        self.total_cost = round(float(total_cost), 2)
        self.gross_profit = round(float(gross_profit), 2)
        try:
            self.gross_margin = round(float(gross_margin), 3)
        except:
            self.gross_margin = 0.0


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

        if "(Book)" in self.product_vendor or "(book)" in self.product_vendor:
            return "Book"

        if "(In-House)" in self.product_vendor or "(in-house)" in self.product_vendor:
            return "In-House"

        if "(Digital)" in self.product_vendor or "(digital)" in self.product_vendor:
            return "Digital"

        return "Unknown"

    def calculate_processing_fee(self):
        return round(float(self.total_sales) * 0.03, 2)

    def calculate_gross_profit(self):
        return self.total_sales - self.calculate_processing_fee() - self.total_cost

    def check_cost(self, dist_type):
        if self.total_cost == 0 and dist_type != "Digital":
            print(
                f"WARNING: Cost is 0 for {self.product_title}. "
                "Please handle this item's cost manually before proceeding manually."
            )

    def detect_anamolies(self):

        if self.net_quantity <= 0 or self.gross_profit < 0:
            return

        if self.product_type == "":
            print(
                f"WARNING: Product Type for {self.product_title} not registered on Shopify. "
                "Please enter information on product in Shopify"
            )
        elif self.product_type.title() not in anamoly_info:
            print(f"WARNING: Product type {self.product_type} for {self.product_title} not registered.")
        elif (
            self.gross_margin < anamoly_info[self.product_type.title()]["Average"] - 0.10
            or self.gross_margin > anamoly_info[self.product_type.title()]["Average"] + 0.10
            ):
            print(
                f"WARNING: {self.product_title} has a gross margin of {self.gross_margin} "
                f"which is outside the expected range of "
                f"{round(anamoly_info[self.product_type.title()]['Average'] - 0.10,3)} to "
                f"{round(anamoly_info[self.product_type.title()]['Average'] + 0.10,3)}. "
                "Please check the item manually."
            )

