# Putting the Pharmacist Scheduler online

The web app (`app.py`) lets anyone you invite upload a schedule workbook, watch it build,
and download the finished copy. These steps host it free on **Streamlit Community Cloud**,
privately. Budget about 20 minutes the first time.

## Before you start: check with IT
The workbook holds staff names and schedules (no patient data). Uploading it to an outside
service may still need your pharmacy leadership's or IT department's OK. The app keeps
nothing: each upload goes to a temporary folder that's deleted as soon as the download is
ready. The hosting service itself still processes the file.

## 0. Try it on your Mac first (optional, 2 minutes)
```bash
cd ~/pharm-scheduler
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```
A browser tab opens at http://localhost:8501. Press Control+C in Terminal to stop it.

## 1. Put the code on GitHub (private)
1. Create a free account at github.com if you don't have one.
2. Click **+** (top right) → **New repository**. Name it `pharm-scheduler`, choose **Private**,
   and click **Create repository**.
3. On the new repository's page, click **uploading an existing file**.
4. In Finder, open your `pharm-scheduler` folder and press **Command + Shift + .** (period) to
   show hidden items. Select everything **except** `.venv`, then drag it onto the GitHub
   page. Check that **.streamlit** is in the list, then click
   **Commit changes**.
   - Must be there: `app.py`, `requirements.txt`, the `scheduler` folder, `.streamlit`.
   - Don't upload the template or any schedule workbook; the repository only needs the code.
     Keep the template internal.

## 2. Create the app
1. Go to **share.streamlit.io** and sign in with your GitHub account. Allow access to your
   private repositories when asked.
2. Click **Create app** → deploy from a GitHub repository.
3. Repository: `<your-username>/pharm-scheduler`, branch `main`, main file `app.py`.
4. Pick an app address (e.g. `pharmacy-schedule`).
5. Open **Advanced settings** and choose **Python 3.12**.
6. Click **Deploy**. The first build takes a few minutes while it installs the solver.

## 3. Make it private and invite people
1. From your workspace at share.streamlit.io, open the app's **⋮** menu → **Settings** →
   **Sharing**.
2. Set **Who can view this app** to **Only specific people**, and add the email addresses of
   the people who should use it. They'll sign in with that email to open it.
3. The free plan allows one private app, which is all this needs.

## Day to day
- **Sleeping:** after a stretch with no visitors the app goes to sleep. The next person sees a
  button to wake it; it's ready again in about a minute.
- **Updating the program:** upload the changed files to the GitHub repository the same way
  (or edit them there). The app rebuilds itself within a minute or two.
- **Search time:** the workbook's Settings sheet sets it (60 s by default). The app's Options
  can override it for one run. On the hosted server, 60 s gives the same quality as on a Mac.
- **Something went wrong?** In the app's **⋮** menu, **Manage app** shows its log.
