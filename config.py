class colors:
    RESET = '\033[0m'
    BOLD = '\033[1m'
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'

extraction_cells = {
    'Clearing': 'D34',
    'Bidding': 'E34',
    'Average Execution Rate': 'I34',
    'Service Income (NTD)': 'J34',
    'Daily Event': ['P30', 'Q30', 'R30']
}

centered_columns = [
    'Clearing', 
    'Bidding', 
    'Average Execution Rate', 
    'Service Income (NTD)'
]

green_columns = [
    'Service Income (NTD)'
]

red_columns = [
    'Daily Event'
]
