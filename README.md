# DRE Automation Tool

This project automates the generation of a Demonstrativo de Resultados do Exercício (DRE) by processing an Excel file (`entrada.xlsx`). The script reads data from various sheets, performs financial calculations, and creates a formatted DRE sheet with key performance indicators.

## Usage

To use this tool, follow these steps:

1. **Prepare the Input File**: Ensure you have an Excel file named `entrada.xlsx` in the same directory as the script. This file must contain the following sheets with the specified columns:
   - `Vendas`: Sales data
   - `Custo_Despesas`: Costs and expenses
   - `Folha`: Payroll information
   - `Investimentos`: Investment details
   - `Financiamento`: Financing data

2. **Configure Parameters**: Open the `parametros.py` file to adjust the script's settings, such as the tax rate and the period for the DRE.

3. **Run the Script**: Execute the `main.py` script to generate the DRE.

   ```bash
   python main.py
   ```

## Input File Structure

The `entrada.xlsx` file must be structured as follows:

- **Vendas**:
  - Column E: Sales values
  - Column F: Dates

- **Custo_Despesas**:
  - Column A: Cost/expense category
  - Column B: Values
  - Column C: Dates

- **Folha**:
  - Column A: Dates
  - Columns C, D, E: Payroll values

- **Investimentos**:
  - Column A: Dates
  - Column B: Descriptions
  - Column C: Values

- **Financiamento**:
  - Data related to financing

## Configuration

The `parametros.py` file allows you to customize the following settings:

- `taxa_imposto`: The tax rate on profit (in percentage).
- `auto_detectar_periodo`: Set to `True` to automatically detect the DRE period from the data, or `False` to use a specific period.
- `periodo_inicio`, `periodo_final`: The start and end period for the DRE (if `auto_detectar_periodo` is `False`).
- `vida_util_ativos`: The useful life of assets for depreciation calculations.

## Dependencies

This project requires the `openpyxl` library to work with Excel files. You can install it using pip:

```bash
pip install openpyxl
```

It is recommended to create a `requirements.txt` file to manage dependencies:

```
openpyxl
```
