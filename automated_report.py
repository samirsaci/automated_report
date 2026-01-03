"""
Automate Operational Reports Distribution

This script automates the creation of warehouse operational reports
and prepares them for email distribution with HTML formatting.
"""

import pandas as pd
import matplotlib.pyplot as plt
import codecs


# Day name mapping
DICT_DAYS = {
    "MON": "Monday",
    "TUE": "Tuesday",
    "WED": "Wednesday",
    "THU": "Thursday",
    "FRI": "Friday",
    "SAT": "Saturday",
    "SUN": "Sunday",
}


def load_and_process_data(week="WEEK-1"):
    """Load and process workload data for a specific week."""
    df_day = pd.read_csv("volumes per day.csv")
    df_plot = df_day[df_day["WEEK"] == week].copy()

    # Calculate KPIs
    df_plot["LINES/ORDER"] = df_plot["LINES"] / df_plot["ORDERS"]
    avg_ratio = "{:.2f} lines/order".format(df_plot["LINES/ORDER"].mean())
    max_ratio = "{:.2f} lines/order".format(df_plot["LINES/ORDER"].max())

    # Maximum Day Lines
    busy_day = DICT_DAYS[df_plot.set_index("DAY")["LINES"].idxmax()]
    max_lines = "{:,} lines".format(df_plot["LINES"].max())

    # Total Workload
    total_lines = "{:,} lines".format(df_plot["LINES"].sum())

    return df_plot, avg_ratio, max_ratio, busy_day, max_lines, total_lines


def create_visual(df_plot, week):
    """Create bar chart visualization."""
    fig, ax = plt.subplots(figsize=(14, 7))
    df_plot.plot.bar(
        figsize=(8, 6),
        edgecolor="black",
        x="DAY",
        y=["ORDERS", "LINES"],
        color=["tab:blue", "tab:red"],
        legend=True,
        ax=ax,
    )
    plt.xlabel("DAY", fontsize=12)
    plt.title("Workload per day (Lines/day)", fontsize=12)
    plt.tight_layout()

    # Save plot
    filename = "visual.png"
    fig.savefig(filename, dpi=fig.dpi)
    plt.close(fig)

    print(f"Visual saved to: {filename}")
    return filename


def prepare_html_report(week, total_lines, busy_day, max_lines, avg_ratio, max_ratio):
    """Prepare HTML report with insights."""
    # Read HTML template
    try:
        with codecs.open("report.html", "r") as f:
            html_string = f.read()

        # Replace placeholders with actual values
        html_string = html_string.replace("WEEK", week)
        html_string = html_string.replace("total_lines", total_lines)
        html_string = html_string.replace("busy_day", busy_day)
        html_string = html_string.replace("max_lines", max_lines)
        html_string = html_string.replace("avg_ratio", avg_ratio)
        html_string = html_string.replace("max_ratio", max_ratio)

        # Save processed HTML
        output_file = f"report_{week}.html"
        with codecs.open(output_file, "w") as f:
            f.write(html_string)

        print(f"HTML report saved to: {output_file}")
        return html_string
    except FileNotFoundError:
        print("report.html template not found. Skipping HTML generation.")
        return None


def display_results(week, total_lines, busy_day, max_lines, avg_ratio, max_ratio):
    """Display report results to console."""
    print("=" * 60)
    print(f"WAREHOUSE WORKLOAD REPORT - {week}")
    print("=" * 60)
    print(f"\nTotal workload: {total_lines}")
    print(f"Busiest day: {busy_day} with {max_lines}")
    print(f"Average lines per order: {avg_ratio}")
    print(f"Maximum lines per order: {max_ratio}")
    print("=" * 60)


def main():
    """Main function to generate the operational report."""
    week = "WEEK-1"

    print("Loading and processing data...")
    df_plot, avg_ratio, max_ratio, busy_day, max_lines, total_lines = (
        load_and_process_data(week)
    )

    print("Creating visualization...")
    create_visual(df_plot, week)

    print("Preparing HTML report...")
    prepare_html_report(week, total_lines, busy_day, max_lines, avg_ratio, max_ratio)

    display_results(week, total_lines, busy_day, max_lines, avg_ratio, max_ratio)

    print("\nDone! Report ready for distribution.")


if __name__ == "__main__":
    main()
