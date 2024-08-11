# Financial Analysis Tool

A Python application that performs financial analysis on transaction data from an Excel file. It provides visualizations of balance trends, top spending categories, and top credit transactions. The application features a dark mode interface and displays summary statistics in a separate window.

## Features

- **Data Visualization**: 
  - Balance over time
  - Top spending categories
  - Top credit transactions

- **Statistics**:
  - Total debits and credits
  - Month-end resume
  - Average transaction amount
  - Highest debit and credit
  - Top spending categories and credits

- **Dark Mode**: The application uses a dark theme for a modern look and better visibility in low-light conditions.

## Requirements

- Python 3.x
- `pandas`
- `matplotlib`
- `tkinter`
- `openpyxl` (for reading Excel files)

You can install the required Python packages using `pip`:

```bash
pip install pandas matplotlib openpyxl
```

## Usage

1. **Prepare Your Excel File**: Make sure your Excel file contains the required data with columns including "Data Movimento", "Valor", "Saldo após movimento", "Tipo", and "Descrição". The sheet name should be specified (default is `'Conta_40244254236'`).

2. **Run the Application**:

   To run the application, use the command line to execute the script with the path to your Excel file:

   ```bash
   python financial_analysis.py path/to/your/excel_file.xlsx
   ```

3. **Interact with the Application**:
   - The main window will display visualizations:
     - **Balance Over Time**: Shows the balance trends over time.
     - **Top Spending Categories**: Displays a bar chart of the top spending categories.
     - **Top Credits**: Displays a bar chart of the top credit transactions.
   - Click the "Show Summary" button to open a new window with detailed summary statistics.
   - Use the "Close" button to exit either the main window or the summary window.

## Code Structure

- `financial_analysis.py`: The main script that loads data, calculates statistics, plots graphs, and manages the GUI.
- `README.md`: This file.

## Contributing

If you would like to contribute to this project, please fork the repository and submit a pull request. Make sure to include a description of your changes and any relevant tests.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For any questions or issues, please open an issue on this GitHub repository
