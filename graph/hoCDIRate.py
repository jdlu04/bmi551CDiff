import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

EXCEL_FILE = "/Users/judy/stony/bmi551/bmi551CDiff/data/cDiff_dummyData2026.xlsx"
SHEET_NAME = "Data"
DATE_COL = "Date"
RATE_COL = "HO Incident Rate per 10000 Patient Days"
TEST_COL = "Total Testing"


def create_dummy_excel(path=EXCEL_FILE, seed=7):
    rng = np.random.default_rng(seed)
    dates = pd.date_range("2014-01-01", "2018-06-01", freq="MS")
    n = len(dates)
    trend = np.interp(np.arange(n), [0, 12, 36, 37, n - 1], [12, 12, 5.5, 6.5, 6.5])
    sd = np.where(np.arange(n) < 12, 2.6, np.where(np.arange(n) < 37, 3.5, 2.0))
    rate = np.clip(trend + rng.normal(0, sd), 1.5, None).round(2)
    testing = np.interp(np.arange(n), [0, 12, 37, n - 1], [300, 290, 170, 160])
    testing = (testing + rng.normal(0, 20, n)).round().astype(int)

    df = pd.DataFrame({DATE_COL: dates, RATE_COL: rate, TEST_COL: testing})
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name=SHEET_NAME, index=False)
        ws = xw.sheets[SHEET_NAME]
        for cell in ws["A"][1:]:
            cell.number_format = "mmm yyyy"
        for col, w in zip("ABC", (14, 42, 16)):
            ws.column_dimensions[col].width = w
    print(f"Created dummy data -> {path}")

def phase_limits(values, method="ci", level=0.95, k=3):
    values = np.asarray(values, float)
    mean, n = values.mean(), len(values)
    sd = values.std(ddof=1) if n > 1 else 0.0
    if method == "sd":
        half = k * sd
    else:
        try:
            from scipy import stats
            crit = stats.t.ppf(0.5 + level / 2, max(n - 1, 1))
        except ImportError:
            crit = 1.96
        half = crit * sd / np.sqrt(n)
    return mean, mean - half, mean + half

class HOControlChart:
    def __init__(self, df, figsize=(14, 8.5)):
        self.df = df.sort_values(DATE_COL).reset_index(drop=True)
        self.fig, self.ax = plt.subplots(figsize=figsize)
        self.phases = [] 
        self.baseline_mean = None
        self._events = []

        ax = self.ax
        ax.plot(self.df[DATE_COL], self.df[RATE_COL], "-o", color="black",
                markersize=7, linewidth=1.4, zorder=5)
        ax.set_ylabel("HO Incident Rate / 10000 Patient Days", fontsize=12)
        ax.set_ylim(-1.5, 20)
        ax.set_yticks(range(0, 21, 5))
        pad = pd.Timedelta(days=60)
        ax.set_xlim(self.df[DATE_COL].min() - pad, self.df[DATE_COL].max() + pad)
        self._testing_row()

    def _testing_row(self, y_top=-0.35, y_bot=-1.0, fontsize=8):
        ax = self.ax
        for i, (d, t) in enumerate(zip(self.df[DATE_COL], self.df[TEST_COL])):
            ax.text(d, y_top if i % 2 == 0 else y_bot, f"{t:.0f}", ha="center",
                    va="center", fontsize=fontsize, zorder=6)
        ax.text(0.02, 0.085, "Total Testing", transform=ax.transAxes,
                fontsize=13, va="bottom")

    def add_phase(self, title, start, end, is_baseline=False,
                  method="ci", level=0.95, k=3, draw_limits=True):
        start, end = pd.Timestamp(start), pd.Timestamp(end)
        sub = self.df[(self.df[DATE_COL] >= start) & (self.df[DATE_COL] <= end)]
        mean, lo, hi = phase_limits(sub[RATE_COL], method, level, k)
        x0, x1 = sub[DATE_COL].min(), sub[DATE_COL].max()

        if draw_limits:
            self.ax.hlines([lo, hi], x0, x1, colors="red", lw=1.8, zorder=3)
            self.ax.hlines(mean, x0, x1, colors="limegreen", lw=1.8, zorder=3)

        if is_baseline:
            self.baseline_mean = mean
            self.ax.hlines(mean, x1, self.ax.get_xlim()[1], colors="limegreen",
                           linestyles="--", lw=1.8, zorder=2)

        title_x = mdates.date2num(start + pd.Timedelta(days=10))
        self.ax.text(title_x, 1.02, title, transform=self.ax.get_xaxis_transform(),
                     ha="left", va="bottom", fontsize=15, fontweight="bold")
        self.phases.append(dict(title=title, start=x0, end=x1, mean=mean, lo=lo, hi=hi))

    def add_phase_divider(self, date):
        self.ax.axvline(pd.Timestamp(date), color="black", lw=1.4, zorder=1)

    def add_event_line(self, date, label):
        date = pd.Timestamp(date)
        self.ax.axvline(date, color="black", linestyle="--", lw=1.6, zorder=2)
        xnum = mdates.date2num(date)
        self.ax.annotate(label, xy=(xnum, 1.0), xycoords=self.ax.get_xaxis_transform(),
                         xytext=(xnum, 1.075), textcoords=self.ax.get_xaxis_transform(),
                         ha="center", va="bottom", fontsize=13,
                         arrowprops=dict(arrowstyle="-|>", color="black", lw=1.6,
                                         shrinkA=0, shrinkB=0),
                         annotation_clip=False)
        self._events.append((date, label))


    def finish(self, tick_dates=None, save_as="ho_chart.png"):
        ax = self.ax
        if tick_dates is None:     
            tick_dates = [d for d in self.df[DATE_COL] if d.month == 1]
            tick_dates.append(self.df[DATE_COL].iloc[-1])
        ax.set_xticks(tick_dates)
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %y"))
        ax.tick_params(axis="x", labelsize=12)
        ax.tick_params(axis="y", labelsize=12)

        for s in ax.spines.values():
            s.set_linewidth(1.6)
        self.fig.subplots_adjust(top=0.88, left=0.07, right=0.98, bottom=0.07)
        self.fig.savefig(save_as, dpi=200)
        print(f"Saved chart -> {save_as}")
        plt.show()

if __name__ == "__main__":
    if not os.path.exists(EXCEL_FILE):
        create_dummy_excel()

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, parse_dates=[DATE_COL])

    chart = HOControlChart(df)

    chart.add_phase("Baseline", "2014-01-01", "2014-12-01", is_baseline=True)
    chart.add_phase("Phase 1",  "2015-01-01", "2016-12-01")
    chart.add_phase("Phase 2", "2017-02-01", "2024-12-01")

    chart.add_phase_divider("2014-12-31")
    chart.add_phase_divider("2017-02-1")
    chart.add_phase_divider("2024-12-01")

    chart.add_event_line("2015-10-01", "a")
    chart.add_event_line("2016-03-01", "b")
    chart.add_event_line("2016-08-01", "c")
    chart.add_event_line("2016-10-01", "d")
    chart.add_event_line("2016-11-01", "e")
    chart.add_event_line("2017-01-01", "f")
    chart.add_event_line("2020-01-01", "g") #two step reflex testing
    chart.add_event_line("2023-10-01", "h") #separate c diff from pcr from full gi panel
    chart.add_event_line("2024-06-01", "i")  #nurse driven protocol
    chart.add_event_line("2025-09-01", "j") #EIA Negative provider alerts
    chart.add_event_line("2025-10-01", "k") #Probiotics removed from power plans
    chart.add_event_line("2026-01-01", "l") #CDI tableau reporting dashboard
    chart.add_event_line("2026-04-01", "m") # Hand washing compliance auditing
    chart.add_event_line("2026-08-01", "n") #14 day reordering rule (c.diff/GI pathogens)
    chart.add_event_line("2026-09-01", "o") #resident diagnosis/treatment education
#    chart.add_phase("Baseline", "2014-01-01", "2014-12-01", is_baseline=True)
#    chart.add_phase("Phase 1",  "2015-01-01", "2016-12-01")
#   chart.add_phase("Phase 2", "2017-02-01", "2024-12-01")

#    chart.add_phase_divider("2014-12-31")
 #   chart.add_phase_divider("2017-02-1")
 #   chart.add_phase_divider("2024-12-01")

    chart.finish(save_as="ho_chart.png")