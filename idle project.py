import pandas as pd
import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
import requests
import matplotlib.pyplot as plt
import ta

class WorkingCapitalAnalysis:
    def __init__(self):
        self.working_capital_data = None
        self.revenue_data = None
        self.share_price_data = None
        self.current_assets_column = "Total Current Assets"
        self.current_liabilities_column = "Total Current Liabilities"
        self.working_capital_ratio = None

    def load_working_capital_data(self, filename):
        try:
            self.working_capital_data = pd.read_excel(filename, header=None)
            messagebox.showinfo("Data Loaded", "Working Capital data loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_revenue_data(self, filename):
        try:
            self.revenue_data = pd.read_excel(filename, header=None)
            messagebox.showinfo("Data Loaded", "Revenue data loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_share_price_data(self, filename):
        try:
            self.share_price_data = pd.read_excel(filename)
            messagebox.showinfo("Data Loaded", "Share Price data loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def analyze_working_capital(self):
        if self.working_capital_data is None:
            messagebox.showwarning("Warning", "No working capital data loaded.")
            return

        try:
            current_assets = float(self.working_capital_data[self.working_capital_data[0] == self.current_assets_column][1].values[0])
            current_liabilities = float(self.working_capital_data[self.working_capital_data[0] == self.current_liabilities_column][1].values[0])

            self.working_capital_ratio = current_assets / current_liabilities

            messagebox.showinfo("Working Capital Analysis", f"Working Capital Ratio: {self.working_capital_ratio}")

            red_flags = []
            if self.working_capital_ratio < 1:
                red_flags.append("Working Capital Ratio is less than 1.")
            if self.working_capital_ratio < 0.5:
                red_flags.append("Working Capital Ratio is less than 0.5.")

            if red_flags:
                messagebox.showwarning("Red Flags in Working Capital", "\n".join(red_flags))
            else:
                messagebox.showinfo("Working Capital Analysis", "No red flags in working capital.")
        except IndexError:
            messagebox.showwarning("Warning", "Column names not found in the working capital data.")
        except ValueError:
            messagebox.showwarning("Warning", "Working capital values are not numeric.")

    def analyze_revenue(self):
        if self.working_capital_ratio is None:
            messagebox.showwarning("Warning", "Working capital analysis is missing. Please perform working capital analysis.")
            return

        user_input = simpledialog.askfloat("Revenue Analysis", "Enter the revenue for the year:")
        if user_input is not None:
            revenue = user_input

            adjusted_revenue = revenue * self.working_capital_ratio
            messagebox.showinfo("Revenue Analysis", f"Adjusted Revenue: {adjusted_revenue}")

            red_flags = []
            if adjusted_revenue < revenue:
                red_flags.append("Adjusted revenue is lower than the actual revenue.")

            if red_flags:
                messagebox.showwarning("Red Flags in Revenue Analysis", "\n".join(red_flags))
            else:
                messagebox.showinfo("Revenue Analysis", "No red flags in revenue.")


    def analyze_share_price(self):
        if self.share_price_data is None:
            messagebox.showwarning("Warning", "No share price data loaded.")
            return

        try:
            df = self.share_price_data.copy()
            df = ta.add_all_ta_features(df, open="Open", high="High", low="Low", close="Close", volume="Volume")

            df['sma_20'] = df['Close'].rolling(window=20).mean()  # 20-day Simple Moving Average
            df['sma_50'] = df['Close'].rolling(window=50).mean()  # 50-day Simple Moving Average
            df['crossover'] = df['sma_20'] - df['sma_50']  # Crossover: sma_20 minus sma_50

            plt.figure(figsize=(12, 6))
            plt.plot(df['Date'], df['Close'], label='Close Price')
            plt.plot(df['Date'], df['sma_20'], label='20-day SMA')
            plt.plot(df['Date'], df['sma_50'], label='50-day SMA')
            plt.xlabel('Date')
            plt.ylabel('Price')
            plt.title('Price Chart with Moving Averages')
            plt.legend()
            plt.show()

            last_crossover = df['crossover'].iloc[-1]
            if last_crossover > 0:
                inference = "Bullish signal: 20-day SMA crossed above 50-day SMA."
            elif last_crossover < 0:
                inference = "Bearish signal: 20-day SMA crossed below 50-day SMA."
            else:
                inference = "No clear signal: 20-day SMA and 50-day SMA are close or overlapping."

            messagebox.showinfo("Share Price Analysis", f"Inference:\n{inference}")
        except KeyError as e:
            messagebox.showerror("Error", f"Column '{e.args[0]}' not found in the share price data.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

class GUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Financial Analysis")
        self.geometry("400x200")
        self.analysis = WorkingCapitalAnalysis()
        self.menu()

    def menu(self):
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Working Capital Data", command=self.load_working_capital_data)
        file_menu.add_command(label="Load Share Price Data", command=self.load_share_price_data)

        analysis_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Analysis", menu=analysis_menu)
        analysis_menu.add_command(label="Perform Working Capital Analysis", command=self.analyze_working_capital)
        analysis_menu.add_command(label="Perform Revenue Analysis", command=self.analyze_revenue)
        analysis_menu.add_command(label="Perform Share Price Analysis", command=self.analyze_share_price)

    def load_working_capital_data(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.analysis.load_working_capital_data(filename)


    def load_share_price_data(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.analysis.load_share_price_data(filename)

    def analyze_working_capital(self):
        self.analysis.analyze_working_capital()

    def analyze_revenue(self):
        self.analysis.analyze_revenue()

    def analyze_share_price(self):
        self.analysis.analyze_share_price()

if __name__ == "__main__":
    gui = GUI()
    gui.mainloop()

