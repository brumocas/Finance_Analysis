import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk
import argparse

# Load and clean the data
def load_data(file_path, sheet_name):
    """Load data from the Excel file and clean it."""
    data = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=4)
    
    def parse_dates(date):
        try:
            return pd.to_datetime(date, format='%d/%m/%Y')
        except ValueError:
            return None

    data['Data Movimento'] = data['Data Movimento'].apply(parse_dates)
    data = data.dropna(subset=['Data Movimento'])

    data['Valor'] = data['Valor'].str.replace('€', '').str.replace('.', '').str.replace(',', '.').astype(float)
    data['Saldo após movimento'] = data['Saldo após movimento'].str.replace('€', '').str.replace('.', '').str.replace(',', '.').astype(float)

    return data

def calculate_summary(data):
    """Calculate and return summary statistics."""
    debits = data[data['Tipo'] == 'Débito']
    credits = data[data['Tipo'] == 'Crédito']

    total_debits = debits['Valor'].sum()
    total_credits = credits['Valor'].sum()

    data['Month'] = data['Data Movimento'].dt.to_period('M')

    month_end_balance = data['Saldo após movimento'].iloc[-1]
    month_start_balance = data['Saldo após movimento'].iloc[0]
    month_resume = month_start_balance - month_end_balance

    top_spending = debits.groupby('Descrição')['Valor'].sum().sort_values(ascending=True).head(10)
    top_credits = credits.groupby('Descrição')['Valor'].sum().sort_values(ascending=False).head(10)

    avg_transaction = data['Valor'].mean()
    highest_debit = debits['Valor'].min()
    highest_credit = credits['Valor'].max()
    
    month_name = data['Data Movimento'].dt.month_name().iloc[0]

    return {
        'Total Debits': f"{total_debits:.2f} €",
        'Total Credits': f"{total_credits:.2f} €",
        'Month Resume': f"{month_resume:.2f} €",
        'Average Transaction Amount': f"{avg_transaction:.2f} €",
        'Highest Debit': f"{highest_debit:.2f} €",
        'Highest Credit': f"{highest_credit:.2f} €",
        'Top Spending Categories': top_spending,
        'Top Credits': top_credits,
        'Month': month_name
    }, debits, credits

def plot_graphs(data, debits, credits, top_spending, top_credits, root):
    """Plot and embed graphs in the tkinter window."""
    plt.style.use('dark_background')
    
    fig, axs = plt.subplots(3, 1, figsize=(12, 14))
    plt.subplots_adjust(hspace=0.5, top=0.95, bottom=0.05)

    # Balance over time
    axs[0].plot(data['Data Movimento'], data['Saldo após movimento'], marker='o', color='cyan')
    axs[0].set_title('Balance Over Time', fontsize=14, color='white')
    axs[0].set_xlabel('Date', fontsize=10, color='white')
    axs[0].set_ylabel('Balance (€)', fontsize=10, color='white')
    axs[0].tick_params(axis='x', rotation=45, labelsize=8, colors='white')
    axs[0].tick_params(axis='y', labelsize=8, colors='white')
    axs[0].grid(True, color='gray')

    # Top spending categories
    axs[1].barh(top_spending.index, top_spending.values, color='salmon')
    axs[1].set_title('Top Spending Categories', fontsize=14, color='white')
    axs[1].set_xlabel('Total Spent (€)', fontsize=10, color='white')
    axs[1].tick_params(axis='x', labelsize=8, colors='white')
    axs[1].tick_params(axis='y', labelsize=8, colors='white')
    axs[1].invert_yaxis()
    axs[1].grid(True, color='gray')

    # Top Credits
    axs[2].barh(top_credits.index, top_credits.values, color='lightgreen')
    axs[2].set_title('Top Credits', fontsize=14, color='white')
    axs[2].set_xlabel('Total Credit (€)', fontsize=10, color='white')
    axs[2].tick_params(axis='x', labelsize=8, colors='white')
    axs[2].tick_params(axis='y', labelsize=8, colors='white')
    axs[2].invert_yaxis()
    axs[2].grid(True, color='gray')

    # Embed the plot in tkinter window
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

def show_summary_gui(summary):
    """Display the summary statistics in a new window."""
    # Create a new top-level window
    stats_window = tk.Toplevel()
    stats_window.title("Summary Statistics")
    stats_window.configure(bg='#2e2e2e')

    # Apply dark theme styles
    style = ttk.Style(stats_window)
    style.theme_use('clam')
    style.configure('TFrame', background='#2e2e2e')
    style.configure('TLabel', background='#2e2e2e', foreground='white')
    style.configure('Dark.TFrame', background='#2e2e2e')
    style.configure('Dark.TButton', background='#444444', foreground='white', borderwidth=1, focusthickness=3, focuscolor='none')
    style.map('Dark.TButton', background=[('active', '#666666')], foreground=[('active', 'white')])

    frame_summary = ttk.Frame(stats_window)
    frame_summary.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    canvas = tk.Canvas(frame_summary, bg='#2e2e2e', highlightthickness=0)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(frame_summary, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Create a frame inside the canvas to hold summary content
    summary_frame = ttk.Frame(canvas, style='Dark.TFrame')
    canvas.create_window((0, 0), window=summary_frame, anchor='nw')
    summary_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Add labels for summary statistics
    label_font = ('Helvetica', 10, 'bold')
    value_font = ('Helvetica', 10)

    ttk.Label(summary_frame, text="Summary Statistics", font=('Helvetica', 14, 'bold'), background='#2e2e2e', foreground='white').pack(anchor='w', pady=5)
    ttk.Label(summary_frame, text=f"Month: {summary['Month']}", font=label_font, background='#2e2e2e', foreground='white').pack(anchor='w')
    ttk.Label(summary_frame, text="Total Debits:", font=label_font, background='#2e2e2e', foreground='white').pack(anchor='w')
    ttk.Label(summary_frame, text=f"{summary['Total Debits']}", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    ttk.Label(summary_frame, text="Total Credits:", font=label_font, background='#2e2e2e', foreground='white').pack(anchor='w')
    ttk.Label(summary_frame, text=f"{summary['Total Credits']}", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    ttk.Label(summary_frame, text="Month Resume:", font=label_font, background='#2e2e2e', foreground='white').pack(anchor='w')
    ttk.Label(summary_frame, text=f"{summary['Month Resume']}", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    ttk.Label(summary_frame, text="Average Transaction Amount:", font=label_font, background='#2e2e2e', foreground='white').pack(anchor='w')
    ttk.Label(summary_frame, text=f"{summary['Average Transaction Amount']}", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    ttk.Label(summary_frame, text="Highest Debit:", font=label_font, background='#2e2e2e', foreground='white').pack(anchor='w')
    ttk.Label(summary_frame, text=f"{summary['Highest Debit']}", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    ttk.Label(summary_frame, text="Highest Credit:", font=label_font, background='#2e2e2e', foreground='white').pack(anchor='w')
    ttk.Label(summary_frame, text=f"{summary['Highest Credit']}", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    # Create a frame for top spending categories
    top_spending_frame = ttk.Frame(summary_frame, style='Dark.TFrame')
    top_spending_frame.pack(fill=tk.X, padx=10, pady=10)

    ttk.Label(top_spending_frame, text="Top Spending Categories:", font=('Helvetica', 12, 'bold'), background='#2e2e2e', foreground='white').pack(anchor='w')

    for description, value in summary['Top Spending Categories'].items():
        ttk.Label(top_spending_frame, text=f"{description}: {value:.2f} €", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    # Create a frame for top credits
    top_credits_frame = ttk.Frame(summary_frame, style='Dark.TFrame')
    top_credits_frame.pack(fill=tk.X, padx=10, pady=10)

    ttk.Label(top_credits_frame, text="Top Credits:", font=('Helvetica', 12, 'bold'), background='#2e2e2e', foreground='white').pack(anchor='w')

    for description, value in summary['Top Credits'].items():
        ttk.Label(top_credits_frame, text=f"{description}: {value:.2f} €", font=value_font, background='#2e2e2e', foreground='white').pack(anchor='w', padx=10)

    # Add a button to close the statistics window
    ttk.Button(stats_window, text="Close", style='Dark.TButton', command=stats_window.destroy).pack(padx=10, pady=10)

def main(file_path):
    sheet_name = 'Conta_40244254236'

    data = load_data(file_path, sheet_name)
    summary, debits, credits = calculate_summary(data)
    
    # Set up the tkinter root window
    root = tk.Tk()
    root.title("Financial Analysis")
    root.configure(bg='#2e2e2e')

    # Apply dark theme styles
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TFrame', background='#2e2e2e')
    style.configure('TLabel', background='#2e2e2e', foreground='white')
    style.configure('Dark.TFrame', background='#2e2e2e')
    style.configure('Dark.TButton', background='#444444', foreground='white', borderwidth=1, focusthickness=3, focuscolor='none')
    style.map('Dark.TButton', background=[('active', '#666666')], foreground=[('active', 'white')])

    # Show the summary statistics in a new window
    ttk.Button(root, text="Show Summary", style='Dark.TButton', command=lambda: show_summary_gui(summary)).pack(padx=10, pady=10)

    # Plot the graphs in the tkinter window
    plot_graphs(data, debits, credits, summary['Top Spending Categories'], summary['Top Credits'], root)

    # Add a button to close the main GUI
    ttk.Button(root, text="Close", style='Dark.TButton', command=root.destroy).pack(padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process an Excel file for financial analysis.')
    parser.add_argument('file_path', type=str, help='Path to the Excel file')
    args = parser.parse_args()
    
    main(args.file_path)
