# EICH Production Planning

## Run locally
1. Install Python 3.11 or 3.12.
2. Open a terminal in this folder.
3. Run `python -m pip install -r requirements.txt`.
4. Run `streamlit run app.py`.
5. Open the local URL printed in the terminal.

## GitHub / Streamlit deployment
Upload `app.py`, `requirements.txt`, and this README to the repository. In Streamlit Community Cloud, select `app.py` as the app entry point.

## Workbook format
The first worksheet must be Orders and the second must be Stock. Import `.xlsx` only. The app uses local `data/` storage by default; local files may not persist after a hosted instance restarts unless persistent storage is configured.

## Password
Default password: `Eich123`. Set an `APP_PASSWORD` environment variable or Streamlit secret for production use.
