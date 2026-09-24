# Deploy to Render and get a public URL

This application is ready to deploy as a Render Web Service.

## Easiest method: GitHub + Render

1. Create a new private GitHub repository, for example `mdv-schedule-converter`.
2. Upload all files from this folder to the repository root. Keep `templates/MDV_Schedule_Template.xlsx` in the templates folder.
3. Sign in to Render (https://render.com) and choose **New > Web Service**.
4. Connect GitHub and select your private repository.
5. Render can read `render.yaml`. If you configure manually, use:
   - Runtime: Python 3
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 120 app:app`
6. Click **Create Web Service** / **Deploy**.
7. When deployment finishes, Render assigns an HTTPS URL similar to:
   `https://mdv-schedule-converter.onrender.com`
8. Open that URL on any computer or phone. No local installation is required for users.

## Custom domain (optional)

In Render, open the service > Settings > Custom Domains and add a domain/subdomain you own, for example `schedule.example.com`. Follow the DNS instructions Render displays. Render handles TLS/HTTPS.

## Important privacy note

A public URL means anyone who knows the URL can reach the upload page unless access control is added. For operational schedules, use a private GitHub repository and add authentication before wider deployment. Uploaded/output files are stored only in the service's ephemeral filesystem, but the current application does not yet enforce user login.

## Railway alternative

Railway can also deploy this repository. Create a Railway project from the GitHub repository. Set the start command to the same Gunicorn command if needed. After deployment, open Service > Settings > Networking > Generate Domain. Railway will create a public `*.up.railway.app` URL.
