# Python script to generate a dynamic graph for RVOL20/GEX, Z-Score, and SPX Close price over time

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
import matplotlib.dates as mdates
import numpy as np

def plot_dynamic_graph(file_path):
    # Load the data
    df = pd.read_excel(file_path)

    # Convert 'Date' column to datetime
    df['Date'] = pd.to_datetime(df['Date'])

    # Set 'Date' as the index
    df.set_index('Date', inplace=True)

    # Calculate rolling 60-trading-day z-score for RVOL20/GEX
    def custom_zscore(x):
        if len(x) < 2:
            return np.nan
        return (x.iloc[-1] - np.nanmean(x)) / np.nanstd(x, ddof=1)

    # Use a 60 trading day window (approximately 3 months)
    df['zscore_rvol20_gex'] = df['gex/rvol20'].rolling(window='60D', min_periods=2).apply(custom_zscore)

    # Create the figure and axis objects
    fig, ax1 = plt.subplots(figsize=(12, 6))

    # Plot RVOL20/GEX
    ax1.plot(df.index, df['gex/rvol20'], label='RVOL20/GEX', color='blue')
    ax1.set_xlabel('Date')
    ax1.set_ylabel('RVOL20/GEX', color='blue')
    ax1.tick_params(axis='y', labelcolor='blue')

    # Create a second y-axis for Z-Score
    ax3 = ax1.twinx()
    ax3.spines['right'].set_position(('outward', 60))  # Offset the second y-axis

    # Plot Z-Score on the second y-axis
    ax3.plot(df.index, df['zscore_rvol20_gex'], label='Z-Score (60 trading days)', color='green', linestyle='--')
    ax3.set_ylabel('Z-Score', color='green')
    ax3.tick_params(axis='y', labelcolor='green')

    # Create a third y-axis for SPX Close price
    ax2 = ax1.twinx()
    ax2.spines['right'].set_position(('outward', 120))  # Offset the third y-axis

    # Plot SPX Close price on the third y-axis
    ax2.plot(df.index, df['SPX Close price'], label='SPX Close Price', color='red')
    ax2.set_ylabel('SPX Close Price', color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    # Set title and adjust layout
    plt.title('RVOL20/GEX, Z-Score (60 trading days), and SPX Close Price Over Time')
    fig.tight_layout()

    # Format x-axis to show dates nicely
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax1.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    plt.gcf().autofmt_xdate()  # Rotation

    # Add legend
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    lines3, labels3 = ax3.get_legend_handles_labels()
    ax1.legend(lines1 + lines2 + lines3, labels1 + labels2 + labels3, loc='upper left')

    # Show grid
    ax1.grid(True, which='both', linestyle='--', linewidth=0.5)

    # Save the plot as a PNG file
    plt.savefig('dynamic_graph_with_zscore.png', dpi=300, bbox_inches='tight')

    # Show the plot
    plt.show()

if __name__ == "__main__":
    plot_dynamic_graph('rvol.study.xlsx')
