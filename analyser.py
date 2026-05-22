from analyser.main_analyser import MainAnalyzer

def analyze_file(filepath: str, customer: str, error_folder: str, unique_id: str):
    """Entry point for the analyzer functionality."""
    analyzer = MainAnalyzer(config_folder="path_to_config_folder")
    return analyzer.run_analyser(filepath, customer, error_folder, unique_id)


analyze_file("V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\aut-22-0684-20240604100830.zip", "SAGE", "V:\\FOR_BREAKDOWN\\ERROR", "sfjskfusikk999kfsdf")
