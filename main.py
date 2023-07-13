# prod
from modules.estimate_data_collector import EstimateBaseTemplate
from modules.estimate_data_writer import NeosintezTemplate

# dev
from pprint import pprint
import os


if __name__ == "__main__":
    # input_folder_path = 'input'
    # for file in os.listdir(input_folder_path):
    #     file_path = os.path.join(input_folder_path, file)
        
    #     current_estimate = EstimateBaseTemplate(file_path)
    #     current_estimate.read_rows()
        
    #     NeosintezTemplate(current_estimate).export()
        
  
    # current_estimate = EstimateBaseTemplate("input/80633-П-06-10-01.ЛСР-5100-102-ТХ - Ресурсная смета (полная форма)1.xlsx")
    # current_estimate = EstimateBaseTemplate("input/80633-Р-02-12-16-2440-ТМ.xlsx")
    # current_estimate = EstimateBaseTemplate("input/03-01-10.xlsx")
    current_estimate = EstimateBaseTemplate("input/80633-П-06-07-08.xlsx")
    # current_estimate = EstimateBaseTemplate("input/03-01-10 _test.xlsx")
    current_estimate.read_rows()
    # print(*current_estimate.rows, sep="\n")
    
    NeosintezTemplate(current_estimate).export()
