# Getting your team a real link - step by step

No coding or command-line tools needed beyond the git commands below (which you've already got working). This takes about 20-25 minutes the first time.

**Already deployed once already?** Skip straight to **Part 2 (MongoDB Atlas)** below, then come back to Part 3 to add the connection string to your *existing* Render service and push the updated code - you don't need to create a new GitHub repo or a new Render service.

## Part 1 - Put the code on GitHub

1. Go to **github.com** and click **Sign up** (use your email - it's free). Verify your email.
2. Once logged in, click the **+** icon top-right → **New repository**.
3. Name it `padmasri-travels-app`. Leave it **Public** (or Private if you prefer - either works with Render's free tier). Click **Create repository**.
4. On the new repo's page, click **uploading an existing file** (or **Add file → Upload files**).
5. Open the `backend-app` folder I've given you, select **all the files and folders inside it** (`server.js`, `package.json`, the `data` folder, the `public` folder) and drag them into the GitHub upload box.
6. Scroll down, click **Commit changes**.

From now on, whenever I hand you updated files, push them from your terminal instead of dragging and dropping:
```bash
cd ~/Downloads/PadmasriTravels_Supervisor_app
git add .
git commit -m "Update"
git push
```
Render redeploys automatically every time you push.

## Part 2 - Create a free MongoDB database (this is what makes data actually stick)

Without this, Render wipes your data every time it redeploys or restarts - which is the problem you just ran into. This fixes it permanently, for free.

1. Go to **mongodb.com/cloud/atlas/register** and sign up (free, use your email).
2. When asked to create a deployment, choose the **free "M0" tier** (sometimes labeled "Free" or "Shared"). Pick any cloud provider/region - AWS Mumbai (ap-south-1) is closest to India if offered. Click **Create**.
3. It'll ask you to create a database user - set a **username and password** (write these down, you'll need them in a moment). Click **Create User**.
4. Under **Network Access** (left sidebar), click **Add IP Address** → **Allow Access from Anywhere** (`0.0.0.0/0`). This is required because Render's servers don't have a fixed address. Click **Confirm**.
   - This is safe here because the database still requires the username/password from step 3 to do anything - "network access" just controls who's *allowed to try* to connect, not who gets in.
5. Go to **Database** (left sidebar) → click **Connect** on your cluster → **Drivers** → copy the connection string. It looks like:
   ```
   mongodb+srv://yourusername:<password>@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority
   ```
6. Replace `<password>` in that string with the actual password from step 3. Save this full string somewhere - you'll paste it into Render next.

## Part 3 - Deploy on Render

**If this is your first deploy:**
1. Go to **render.com** → **Get Started** → sign up (or "Sign up with GitHub").
2. **New +** → **Web Service** → select the `padmasri-travels-app` repo.
3. Fill in:
   - **Name**: `padmasri-travels`
   - **Region**: Singapore
   - **Branch**: `main`
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Instance Type**: **Free**
4. Before clicking Create, scroll to **Environment Variables** → **Add Environment Variable**:
   - Key: `MONGODB_URI`
   - Value: the full connection string from Part 2, step 6
5. Click **Create Web Service** and wait 2-3 minutes for it to say **Live**.

**If you already have a Render service running (your situation right now):**
1. Push the updated code first (see the git commands at the end of Part 1 - I've updated `server.js` and `package.json`).
2. On your service's page in Render, go to **Environment** (left sidebar) → **Add Environment Variable**:
   - Key: `MONGODB_URI`
   - Value: the full connection string from Part 2, step 6
3. Click **Save Changes** - Render will automatically redeploy with the new code and the database connection.
4. Check the **Logs** tab and look for `Storage: connected to MongoDB` - that confirms it worked. If instead you see `MONGODB_URI not set`, double check the environment variable was saved.

Your link (`https://padmasri-travels.onrender.com` or similar) is what you send your team. From now on, your data survives redeploys, restarts, and sleep/wake cycles.

## Important things to know about the free tier

- **Render sleeps.** After 15 minutes of no visitors it puts the app to sleep; the next visit takes 30-60 seconds to wake up. Normal, not a bug.
- **MongoDB Atlas free tier (M0) gives you 512MB of storage**, which is plenty for records, users, vehicles, and text data, and holds a meaningful number of photos too (they're compressed on the phone before upload). If you eventually outgrow it, upgrading is a paint-free plan change in Atlas, not a rebuild.
- **This is genuinely shared, persistent data** now - not a preview, not something that resets. Treat PINs and access accordingly.

## Backups (data kept for 30 days)

Atlas's free tier doesn't include automatic cloud backups (that's a paid-tier feature), so the app now takes its own: once a day it saves a full snapshot of all your data into a separate part of the same database, and automatically deletes anything older than 30 days.

- As the **Owner**, scroll to the bottom of the **Dashboard** to see the **Backups** card - a list of the last 30 days of snapshots.
- If something goes badly wrong (a bad bulk edit, an accidental deletion), click **Restore this** next to the snapshot from before it happened. This replaces all current data with that snapshot - use it only to undo a real mistake, since anything entered after that snapshot's time is lost.
- This protects against mistakes made *through the app*. It does not protect against Atlas itself going down, which Atlas already guards against on its own infrastructure.

There's also an optional `backup/` folder in the project with a script that saves a copy of your data as plain JSON files onto your own computer, if you'd like an extra offline copy - `backup/config.json` already has your connection string in it (not uploaded to GitHub). Run it anytime from a terminal with `node backup/backup.js`.

## Adding your real team

Once the link is live:
1. Log in as the Owner (Vamshi Reddy, PIN `9999` - **change this PIN first**, via **People → Edit**).
2. Go to **People** → add each of your real supervisors, area supervisors, HR person, ops manager and data team members with their own PIN, or **Edit** the seeded demo accounts (Ravi, Suresh, Priya, Anil/Hemanth, Sneha, Divya) to become real people.
3. Go to **Assignments** → add your real sites, drivers, and vehicles, and assign each vehicle to a site, supervisor, driver and client.
4. Give each person their own PIN privately - shared logins defeat the audit log.

## If anything goes wrong

Send me a screenshot of the error (Render's **Logs** tab, or the MongoDB Atlas dashboard), and I'll help you fix it from here.
