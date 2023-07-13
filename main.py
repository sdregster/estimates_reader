# prod
from modules.estimate_data_collector import EstimateBaseTemplate
from modules.estimate_data_writer import NeosintezTemplate

# dev
from pprint import pprint


if __name__ == "__main__":
    current_estimate = EstimateBaseTemplate("80633-П-06-07-08.xlsx")
    current_estimate.read_rows()

    print(*current_estimate.rows, sep="\n")
    
    # NeosintezTemplate(current_estimate).export()
