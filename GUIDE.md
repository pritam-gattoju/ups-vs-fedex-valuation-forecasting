# Step-by-step guide

This guide takes you from the downloaded zip to a live dashboard linked from your resume. No coding experience is needed for Parts 1-4. Plan on about an hour the first time.

---

## Part 1 - Put the project on GitHub (10 min)

1. Sign in at **github.com**. Click **+** (top right) > **New repository**.
2. Name it `ups-vs-fedex-valuation-forecasting`, set it to **Public**, tick **Add a README file** (you will replace it), and click **Create repository**.
3. Unzip `ups-vs-fedex-valuation-forecasting.zip` on your computer.
4. In your new repo click **Add file > Upload files**. Open the unzipped folder, select **everything inside it** (not the folder itself), and drag it into the browser window. Folders upload with their contents.
5. Scroll down, type a commit message such as `Initial upload`, and click **Commit changes**.
6. Check that the repo shows the folders `data`, `src`, `model`, `report`, `notebooks`, `outputs`, `scripts`, `tests`, `images`, plus `app.py` and `README.md`.
   - The `.streamlit` folder is hidden on most computers. If it does not upload, that is fine; it only sets the dashboard colors.
7. Your GitHub username (pritam-gattoju) is already filled into every link, so there is nothing to edit.

---

## Part 2 - Run the forecasting notebook in Google Colab (15 min)

The notebook downloads real market data, trains the four models and creates the files the dashboard's **Price forecasting** tab needs.

1. On your repo page, open the README and click the **Open in Colab** badge.
   - If the badge does not work: go to **colab.research.google.com** > **File > Open notebook > GitHub**, paste your repo URL, and pick `notebooks/01_market_data_and_ml_forecasting.ipynb`.
2. Click **Runtime > Run all**. If Colab warns the notebook is not from Google, click **Run anyway**.
3. Wait 5-10 minutes. You will see charts appear. At the end, your browser downloads **`forecast_outputs.zip`** (allow the download if asked).
4. Unzip it. You will have 8 files such as `model_metrics.csv` and `actual_vs_predicted.csv`.
5. On GitHub, click into the **`outputs`** folder, then **Add file > Upload files**. Drag in all 8 files and **Commit changes**.
6. Before you close Colab, read the **model_metrics** table in Step 4. These are your real results; write down which model had the lowest RMSE for UPS and for FedEx. They make good interview talking points.

---

## Part 3 - Deploy the dashboard for free (10 min)

1. Go to **share.streamlit.io** and sign in with your GitHub account.
2. Click **Create app** > **Deploy a public app from GitHub**.
3. Choose your repository, branch `main`, and main file path **`app.py`**.
4. Optional: under **Advanced settings**, choose Python **3.12**.
5. Pick a custom URL such as `ups-fedex-valuation`, then click **Deploy**. The first build takes 2-4 minutes.
6. Click through every tab to check it works. Copy the URL.
7. Back on GitHub, edit `README.md` and replace `_add your Streamlit link here after deploying_` with your link, for example `[ups-fedex-valuation.streamlit.app](https://ups-fedex-valuation.streamlit.app)`.
8. Take a screenshot of the **DCF valuation lab** tab, upload it to the `images` folder as `dashboard.png`, and add this line under the badges in the README: `![Dashboard](images/dashboard.png)`.

> Free Streamlit apps go to sleep after a few days without visitors. Anyone who opens the link can wake it with one click. Open it yourself the day before an interview.

---

## Part 4 - Show it to recruiters (10 min)

**GitHub profile:** go to your profile page > **Customize your pins** > select this repo so it appears first.

**Resume - Projects section:**

> **UPS vs FedEx: Valuation, Strategy & Forecasting** | Excel, Python, Streamlit | github.com/pritam-gattoju/ups-vs-fedex-valuation-forecasting
> - Built a 15-sheet, 1,500-formula DCF model from UPS and FedEx 10-Ks (ratio and DuPont analysis, CAPM WACC, reverse DCF, sensitivity and strategy scenarios); a reverse DCF showed UPS's price implied an 8.1% margin vs 9.6% guidance (+25% upside) and FedEx's a 9.5% margin vs 7.0% actual (30% downside).
> - Developed a Python valuation engine with 10,000-run Monte Carlo and trained ARIMA, gradient boosting and LSTM price forecasters on 16 years of daily data; deployed an interactive Streamlit dashboard.

Use one bullet if space is tight. Put your GitHub link in your resume header next to LinkedIn.

**LinkedIn:** Profile > **Add profile section > Recommended > Add featured** > **Link**, and paste the Streamlit dashboard URL.

---

## Part 5 - Run it on your own computer (optional)

1. Install Python 3.12 from python.org (on Windows, tick **Add Python to PATH**).
2. Open a terminal in the project folder and run:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Your browser opens the dashboard at `http://localhost:8501`.

To rebuild the Excel model and report after editing a CSV in `data/`:
```
pip install -r requirements-dev.txt
python scripts/build_excel_model.py
python scripts/build_report.py
```
Then open the workbook in Excel and save it once, so it stores the recalculated values, and run `python -m pytest tests -q` to confirm Excel and Python still agree.

---

## Part 6 - Know it cold for interviews

Be ready to explain these five things in your own words:

1. **Why the two stocks diverged:** UPS's price implied an 8.1% operating margin, below its own 9.6% guidance, because investors doubted the Amazon glide-down plan. FedEx's implied 9.5%, above its 7.0% today, because investors had priced in Network 2.0 savings and the Freight spin-off.
2. **How WACC was built:** CAPM cost of equity = risk-free rate + beta x equity risk premium; cost of debt = interest expense / average debt, after tax; weights at market value.
3. **Why terminal value dominates:** it is 73-78% of enterprise value, which is why WACC and long-run growth are the most powerful value drivers.
4. **What DuPont revealed:** UPS's ROE (34%) is more than twice FedEx's (14%) because it wins on three of the five factors: a higher operating margin (the biggest gap), faster asset turnover, and more leverage (assets are 4.5x equity after years of buybacks and dividends). Part of its ROE edge is borrowed, not earned.
5. **What the ML models showed:** daily prices behave close to a random walk, so beating the naive forecast is hard. That is why the valuation relies on fundamentals, and saying so shows judgment.

Open `model/UPS_vs_FedEx_Valuation_Model.xlsx` and change a yellow assumption (for example, FedEx operating margin to 9.5%) to watch the value per share approach the share price. Practice doing this live.
