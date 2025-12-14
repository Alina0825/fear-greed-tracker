import tkinter as tk
from tkinter import ttk, messagebox
import requests
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib
import threading
import mplcursors
import time

# Set font configuration
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['axes.unicode_minus'] = False


class FearGreedTracker:
    """CNN Fear & Greed Index Tracker"""

    def __init__(self, root):
        self.root = root
        self.root.title("CNN Fear & Greed Index Tracker")
        self.root.geometry("1000x700")

        self.base_url = 'https://production.dataviz.cnn.io/index/fearandgreed/graphdata'
        self.excel_file = 'fear_greed_index.xlsx'
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json',
            'Referer': 'https://edition.cnn.com/'
        }

        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        """Setup user interface"""
        # Top information panel
        info_frame = ttk.Frame(self.root, padding="10")
        info_frame.pack(fill=tk.X)

        # Current index display
        self.score_label = tk.Label(info_frame, text="--", font=("Arial", 48, "bold"), fg="#333")
        self.score_label.pack()

        self.rating_label = tk.Label(info_frame, text="Loading...", font=("Arial", 20), fg="#666")
        self.rating_label.pack()

        self.time_label = tk.Label(info_frame, text="", font=("Arial", 10), fg="#999")
        self.time_label.pack()

        # Change information
        self.change_label = tk.Label(info_frame, text="", font=("Arial", 12))
        self.change_label.pack(pady=5)

        # Button panel
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X)

        self.refresh_btn = ttk.Button(button_frame, text="🔄 Refresh", command=self.refresh_data)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="📊 View Excel", command=self.open_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="📈 30 Days", command=lambda: self.update_chart(30)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="📈 90 Days", command=lambda: self.update_chart(90)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="📈 1 Year", command=lambda: self.update_chart(365)).pack(side=tk.LEFT, padx=5)

        # Status bar
        self.status_label = tk.Label(self.root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        # Chart panel
        chart_frame = ttk.Frame(self.root, padding="10")
        chart_frame.pack(fill=tk.BOTH, expand=True)

        self.figure = Figure(figsize=(10, 5), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, chart_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def load_data(self):
        """Load historical data"""
        if Path(self.excel_file).exists():
            try:
                self.df = pd.read_excel(self.excel_file, sheet_name='Data')
                self.df['Date'] = pd.to_datetime(self.df['Date'])
                if self.df['Date'].dt.tz is not None:
                    self.df['Date'] = self.df['Date'].dt.tz_localize(None)
                self.status_label.config(text=f"Loaded {len(self.df)} historical records")
            except Exception as e:
                self.df = pd.DataFrame()
                self.status_label.config(text=f"Load failed: {e}")
        else:
            self.df = pd.DataFrame()
            self.status_label.config(text="No historical data found. Fetching...")

        # Check if we need to fetch historical data
        if self.df.empty or len(self.df) < 30:
            self.root.after(1000, self.fetch_initial_historical_data)
        else:
            self.root.after(1000, self.auto_initial_refresh)

    def fetch_initial_historical_data(self):
        """Fetch historical data on first run"""

        def task():
            self.status_label.config(text="Fetching historical data (1 year)...")
            try:
                time.sleep(1)
                # Get 1 year of historical data
                start_date = (datetime.now() - timedelta(days=365)).strftime('%Y-%m-%d')
                url = f"{self.base_url}/{start_date}"

                response = requests.get(url, headers=self.headers, timeout=15)
                response.raise_for_status()
                data = response.json()

                historical = data['fear_and_greed_historical']['data']
                records = []

                for item in historical:
                    timestamp_ms = item['x']
                    score = item['y']
                    rating = item['rating']
                    dt = datetime.fromtimestamp(timestamp_ms / 1000)

                    records.append({
                        'Date': dt,
                        'Score': score,
                        'Rating': rating.title(),
                        'Prev Close': None,
                        '1 Week Ago': None,
                        '1 Month Ago': None
                    })

                self.df = pd.DataFrame(records).sort_values('Date', ascending=False).reset_index(drop=True)
                self.save_to_excel()

                self.root.after(0, lambda: self.status_label.config(
                    text=f"✓ Loaded {len(self.df)} historical records"))
                self.root.after(0, lambda: self.update_chart(365))

                # Now refresh current data
                self.root.after(1000, self.auto_initial_refresh)

            except Exception as e:
                self.root.after(0, lambda: self.status_label.config(
                    text=f"Failed to fetch historical data: {e}"))
                print(f"Error fetching historical data: {e}")

        threading.Thread(target=task, daemon=True).start()

    def auto_initial_refresh(self):
        """Auto refresh on startup"""
        self.status_label.config(text="Auto-refreshing latest data...")
        self.refresh_data()

    def fetch_current_data(self):
        """Fetch current Fear & Greed Index data"""
        try:
            time.sleep(1)
            response = requests.get(self.base_url, headers=self.headers, timeout=15)
            response.raise_for_status()
            data = response.json()

            current = data['fear_and_greed']
            dt = datetime.fromisoformat(current['timestamp'].replace('Z', '+00:00')).replace(tzinfo=None)

            return {
                'date': dt,
                'score': current['score'],
                'rating': current['rating'],
                'prev_close': current.get('previous_close'),
                'prev_week': current.get('previous_1_week'),
                'prev_month': current.get('previous_1_month')
            }
        except Exception as e:
            print(f"Error fetching current data: {e}")
            return None

    def get_rating_color(self, score):
        """Get color based on score"""
        if score < 25:
            return '#ea3943'  # Red - Extreme Fear
        elif score < 45:
            return '#f1ad5d'  # Orange - Fear
        elif score < 55:
            return '#f1d55d'  # Yellow - Neutral
        elif score < 75:
            return '#93d66f'  # Light Green - Greed
        else:
            return '#16c784'  # Green - Extreme Greed

    def refresh_data(self):
        """Refresh data from API"""

        def task():
            self.refresh_btn.config(state='disabled', text="⏳ Refreshing...")
            self.status_label.config(text="Fetching latest data...")

            current = self.fetch_current_data()
            if current:
                # Update dataframe
                current_date = current['date'].replace(tzinfo=None)

                new_row = {
                    'Date': current_date,
                    'Score': current['score'],
                    'Rating': current['rating'].title(),
                    'Prev Close': current.get('prev_close'),
                    '1 Week Ago': current.get('prev_week'),
                    '1 Month Ago': current.get('prev_month')
                }

                if not self.df.empty:
                    # Ensure Date column has no timezone
                    if self.df['Date'].dt.tz is not None:
                        self.df['Date'] = self.df['Date'].dt.tz_localize(None)

                    existing = self.df[self.df['Date'].dt.date == current_date.date()]
                    if not existing.empty:
                        idx = existing.index[0]
                        for key, value in new_row.items():
                            self.df.at[idx, key] = value
                    else:
                        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
                else:
                    self.df = pd.DataFrame([new_row])

                self.df = self.df.sort_values('Date', ascending=False).reset_index(drop=True)
                self.save_to_excel()

                # Update UI
                self.root.after(0, lambda: self.update_display(current))
                self.root.after(0, lambda: self.update_chart(30))
                self.root.after(0, lambda: self.status_label.config(
                    text=f"✓ Updated successfully - {datetime.now().strftime('%H:%M:%S')}"))
            else:
                self.root.after(0, lambda: self.status_label.config(text="✗ Failed to fetch data. Check network."))
                self.root.after(0, lambda: messagebox.showerror("Error",
                                                                "Unable to fetch data. Check network connection."))

            self.root.after(0, lambda: self.refresh_btn.config(state='normal', text="🔄 Refresh"))

        # Run in separate thread to avoid UI freezing
        threading.Thread(target=task, daemon=True).start()

    def update_display(self, current):
        """Update display with current data"""
        score = current['score']
        rating = current['rating'].title()
        color = self.get_rating_color(score)

        self.score_label.config(text=f"{score:.1f}", fg=color)
        self.rating_label.config(text=rating, fg=color)
        self.time_label.config(text=f"Updated: {current['date'].strftime('%Y-%m-%d %H:%M')}")

        # Display change
        if current.get('prev_close'):
            change = score - current['prev_close']
            arrow = "↑" if change > 0 else "↓" if change < 0 else "→"
            change_color = "green" if change > 0 else "red" if change < 0 else "gray"
            self.change_label.config(
                text=f"Daily Change: {arrow} {abs(change):.2f}",
                fg=change_color
            )

    def update_chart(self, days):
        """Update chart with specified timeframe"""
        if self.df.empty:
            return

        self.ax.clear()

        # Get recent N days data
        recent_df = self.df.head(days).sort_values('Date')

        # Adjust marker size based on number of days
        if days <= 30:
            marker_size = 5
            line_width = 2.5
            show_markers = True
        elif days <= 90:
            marker_size = 3
            line_width = 2
            show_markers = True
        else:  # 365 days
            marker_size = 0  # No markers for 1 year view
            line_width = 2
            show_markers = False

        # Plot line chart
        if show_markers:
            line = self.ax.plot(recent_df['Date'], recent_df['Score'],
                                linewidth=line_width, marker='o', markersize=marker_size,
                                color='#1f77b4', markerfacecolor='#1f77b4',
                                markeredgewidth=0, zorder=3)
        else:
            line = self.ax.plot(recent_df['Date'], recent_df['Score'],
                                linewidth=line_width, color='#1f77b4', zorder=3)

        # Add fear/greed zones
        self.ax.axhspan(0, 25, alpha=0.15, color='red', label='Extreme Fear')
        self.ax.axhspan(25, 45, alpha=0.15, color='orange', label='Fear')
        self.ax.axhspan(45, 55, alpha=0.15, color='yellow', label='Neutral')
        self.ax.axhspan(55, 75, alpha=0.15, color='lightgreen', label='Greed')
        self.ax.axhspan(75, 100, alpha=0.15, color='green', label='Extreme Greed')

        # Add red dashed line for today's value
        if not self.df.empty:
            latest_date = self.df.iloc[0]['Date']
            latest_score = self.df.iloc[0]['Score']

            # Vertical dashed line at today's date
            self.ax.axvline(x=latest_date, color='red', linestyle='--',
                            linewidth=2, alpha=0.7, label='Today', zorder=2)

            # Horizontal dashed line at today's score
            self.ax.axhline(y=latest_score, color='red', linestyle='--',
                            linewidth=1.5, alpha=0.5, zorder=2)

        self.ax.set_xlabel('Date', fontsize=11)
        self.ax.set_ylabel('Index Score', fontsize=11)
        self.ax.set_title(f'CNN Fear & Greed Index Trend (Last {days} Days)', fontsize=13, fontweight='bold')
        self.ax.legend(loc='upper left', fontsize=9)
        self.ax.grid(True, alpha=0.3, linestyle='--')
        self.ax.set_ylim(0, 100)

        # Rotate x-axis labels
        self.figure.autofmt_xdate()

        # Add interactive tooltips
        try:
            # hover=2 means: show on hover, disappear when mouse leaves
            cursor = mplcursors.cursor(line, hover=2)

            @cursor.connect("add")
            def on_add(sel):
                # Get data point index
                index = int(sel.index)
                date = recent_df.iloc[index]['Date']
                score = recent_df.iloc[index]['Score']
                rating = recent_df.iloc[index]['Rating']

                # Set annotation text
                sel.annotation.set_text(
                    f"Date: {date.strftime('%Y-%m-%d')}\n"
                    f"Score: {score:.2f}\n"
                    f"Rating: {rating}"
                )
                sel.annotation.get_bbox_patch().set(boxstyle='round,pad=0.5',
                                                    facecolor='#f0f0f0',
                                                    alpha=0.95,
                                                    edgecolor='#666666',
                                                    linewidth=1.5)
                sel.annotation.set_fontsize(10)
                sel.annotation.set_color('#333333')

        except Exception as e:
            # Silent fail if mplcursors unavailable
            print(f"Interactive tooltip failed: {e}")

        self.canvas.draw()

    def save_to_excel(self):
        """Save data to Excel"""
        try:
            with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                self.df.to_excel(writer, sheet_name='Data', index=False)

                if not self.df.empty:
                    summary_data = {
                        'Metric': [
                            'Current Score',
                            'Current Rating',
                            'Last Updated',
                            '7-Day Average',
                            '30-Day Average',
                            'Historical High',
                            'Historical Low',
                            'Total Records'
                        ],
                        'Value': [
                            f"{self.df.iloc[0]['Score']:.2f}",
                            self.df.iloc[0]['Rating'],
                            self.df.iloc[0]['Date'].strftime('%Y-%m-%d %H:%M'),
                            f"{self.df.head(7)['Score'].mean():.2f}",
                            f"{self.df.head(30)['Score'].mean():.2f}",
                            f"{self.df['Score'].max():.2f}",
                            f"{self.df['Score'].min():.2f}",
                            len(self.df)
                        ]
                    }
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)
        except Exception as e:
            print(f"Failed to save Excel: {e}")

    def open_excel(self):
        """Open Excel file"""
        if Path(self.excel_file).exists():
            import os
            os.startfile(self.excel_file)
        else:
            messagebox.showinfo("Info", "Excel file not found. Please refresh data first.")


def main():
    root = tk.Tk()
    app = FearGreedTracker(root)
    root.mainloop()


if __name__ == "__main__":
    main()