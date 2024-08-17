# Python script to generate a dynamic graph for RVOL20/GEX, Z-Score, and SPX Close price over time

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
import matplotlib.dates as mdates
from scipy.stats import zscore

def plot_dynamic_graph(file_path):
    # Load the data
    df = pd.read_excel(file_path)

    # Convert 'Date' column to datetime
    df['Date'] = pd.to_datetime(df['Date'])

    # Calculate rolling 60-day z-score for RVOL20/GEX
    df['zscore_rvol20_gex'] = df['gex/rvol20'].rolling(window=60).apply(lambda x: zscore(x)[-1] if len(x) == 60 else None)

    # Create the figure and axis objects
    fig, ax1 = plt.subplots(figsize=(12, 6))

    # Plot RVOL20/GEX
    ax1.plot(df['Date'], df['gex/rvol20'], label='RVOL20/GEX', color='blue')
    ax1.set_xlabel('Date')
    ax1.set_ylabel('RVOL20/GEX', color='blue')
    ax1.tick_params(axis='y', labelcolor='blue')

    # Plot Z-Score
    ax1.plot(df['Date'], df['zscore_rvol20_gex'], label='Z-Score (60-day)', color='green', linestyle='--')

    # Create a second y-axis
    ax2 = ax1.twinx()

    # Plot SPX Close price on the second y-axis
    ax2.plot(df['Date'], df['SPX Close price'], label='SPX Close Price', color='red')
    ax2.set_ylabel('SPX Close Price', color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    # Set title and adjust layout
    plt.title('RVOL20/GEX, Z-Score (60-day), and SPX Close Price Over Time')
    fig.tight_layout()

    # Format x-axis to show dates nicely
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax1.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    plt.gcf().autofmt_xdate()  # Rotation

    # Add legend
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')

    # Show grid
    ax1.grid(True, which='both', linestyle='--', linewidth=0.5)

    # Save the plot as a PNG file
    plt.savefig('dynamic_graph_with_zscore.png', dpi=300, bbox_inches='tight')

    # Show the plot
    plt.show()

if __name__ == "__main__":
    plot_dynamic_graph('rvol.study.xlsx')
