[README.md](https://github.com/user-attachments/files/31362334/README.md)
# Private Portfolio Allocation System

Files:
- `portfolio_engine.py` — calculation engine (screening, Markowitz, gamma solver, two-fund separation)
- `run_advisor_session.py` — command-line tool for daily desktop use (also used by the web app to build the Excel report)
- `streamlit_app.py` — web app (upload Excel files, enter investor inputs, get results + downloadable report)
- `requirements.txt` — Python packages needed

## 1. Run locally (no internet needed after setup)

1. Install Python 3.10+ if you don't have it: https://www.python.org/downloads/
2. Open a terminal in this folder and install the packages:
   ```
   pip install -r requirements.txt
   ```
3a. **Command-line version** (interactive, asks tenor/amount/preference in the terminal):
   ```
   python run_advisor_session.py
   ```
   Make sure `Portfolio.xlsx` and `Full_Dataset_5_Stocks.xlsx` are in the same folder (or pass paths
   with `--portfolio` / `--dataset`). It prints the report and saves an Excel report file.

3b. **Web app version** (browser UI, upload files each session):
   ```
   streamlit run streamlit_app.py
   ```
   This opens automatically at `http://localhost:8501` in your browser. Only you can see it — it's
   running on your own computer, nothing leaves your machine.

## 2. Put it online for free

Two free options that both work well for a small Streamlit app like this one.

### Option A — Streamlit Community Cloud (simplest, made for this)

1. Create a free GitHub account (github.com) if you don't have one.
2. Create a new **private** repository (e.g. `portfolio-allocator`) and upload these 4 files:
   `streamlit_app.py`, `portfolio_engine.py`, `run_advisor_session.py`, `requirements.txt`.
   (Do **not** upload `Portfolio.xlsx` — you upload that fresh through the app each day instead,
   since it changes daily and contains your live pricing.)
3. Go to https://share.streamlit.io , sign in with GitHub, click "New app", pick your repository
   and `streamlit_app.py` as the entry file, and click Deploy.
4. You'll get a URL like `https://your-app-name.streamlit.app`. Bookmark it — that's your online
   system. Each time you open it, upload the day's `Portfolio.xlsx` and
   `Full_Dataset_5_Stocks.xlsx` in the sidebar, then enter the investor's tenor/amount/preference.
5. To update after any change to the code: just push the change to GitHub — the app redeploys
   automatically.

Free tier limits: the app "sleeps" after a period of no visits and wakes up in a few seconds on
the next visit — fine for a tool you open a few times a day.

### Option B — Hugging Face Spaces (also free, alternative if you'd rather not use GitHub)

1. Create a free account at https://huggingface.co
2. Click "New Space" → give it a name → SDK: **Streamlit** → Hardware: **CPU basic (free)**.
3. Upload the same 4 files through the Space's "Files" tab (or connect a GitHub repo the same way).
4. The Space builds automatically and gives you a URL like
   `https://huggingface.co/spaces/your-username/your-space-name`.

## 3. Privacy note

Both options above are free **public-infrastructure** hosting — the app itself can be made private
(so only people with the link, or only you, can open it), but the investor data you upload each
session (prices, forecast, amounts) is processed on Streamlit's / Hugging Face's servers while the
app runs. For a private wealth-management tool with real client data, if that matters for your
compliance requirements, running it locally (Option in Section 1) or on a paid private server you
control is the safer choice. Free hosting is best treated as a demo / personal-use convenience,
not a client-data system, unless you've checked the provider's terms for your use case.
