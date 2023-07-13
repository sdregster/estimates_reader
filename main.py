import os

from modules.estimate_data_collector import EstimateBaseTemplate
from modules.estimate_data_writer import NeosintezTemplate


if __name__ == "__main__":
    # Построчное чтение всех смет из папки input
    input_folder_path = "input"
    found_estimates = []
    for file in os.listdir(input_folder_path):
        file_path = os.path.join(input_folder_path, file)

        temp_estimate = EstimateBaseTemplate(file_path)
        temp_estimate.read_rows()

        found_estimates.append(temp_estimate)

    # Экспорт в шаблон неосинтеза
    NeosintezTemplate().export(found_estimates)
