# Getting your team a real link - step by step

No coding or command-line tools needed. This takes about 15-20 minutes the first time. You're doing two things: putting the `backend-app` folder onto GitHub (just a file storage website), then pointing Render (a free hosting service) at it.

## Part 1 - Put the code on GitHub

1. Go to **github.com** and click **Sign up** (use your email - it's free). Verify your email.
2. Once logged in, click the **+** icon top-right → **New repository**.
3. Name it `padmasri-travels-app`. Leave it **Public** (or Private if you prefer - either works with Render's free tier). Click **Create repository**.
4. On the new repo's page, click **uploading an existing file** (or **Add file → Upload files**).
5. Open the `backend-app` folder I've given you, select **all the files and folders inside it** (`server.js`, `package.json`, the `data` folder, the `public` folder - everything inside `backend-app`, not the `backend-app` folder itself) and drag them into the GitHub upload box.
   - Make sure the folder structure is preserved (GitHub's drag-and-drop keeps subfolders intact when you drag a folder in most browsers - if it flattens them, let me know and I'll repackage as a zip with instructions).
6. Scroll down, click **Commit changes**.

You now have your code on GitHub. You will come back here later whenever I give you updated files - just re-upload the changed file(s) the same way and click Commit changes; Render redeploys automatically.

## Part 2 - Deploy it on Render (this gives you the real https:// link)

1. Go to **render.com** and click **Get Started** - sign up with the same email, or "Sign up with GitHub" (recommended - it auto-connects).
2. Click **New +** → **Web Service**.
3. Connect your GitHub account if asked, then select the `padmasri-travels-app` repository.
4. Fill in:
   - **Name**: `padmasri-travels` (this becomes part of your link)
   - **Region**: closest to India (Singapore is usually best)
   - **Branch**: `main`
   - **Root Directory**: leave blank
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Instance Type**: **Free**
5. Click **Create Web Service**.
6. Wait 2-3 minutes while it builds. When it says "Live", your link is shown at the top of the page - something like `https://padmasri-travels.onrender.com`.

That's it - that link is what you send to your team. Anyone who opens it on their phone gets the same shared app; whatever one person enters, everyone with the right access can see.

## Important things to know about the free tier

- **It sleeps.** Render's free plan puts the app to sleep after 15 minutes of no visitors, and takes about 30-60 seconds to wake up on the next visit. That's fine for supervisors checking in a few times a day, but the first person each morning will see a short loading delay.
- **Data can reset on redeploy.** This app stores data in a file on Render's disk. On the free plan, that file is wiped whenever you push new code/redeploy (not on every sleep/wake - just on actual redeploys). For real day-to-day use with your full 250+ vehicle fleet, the next step is to attach a persistent disk (a few dollars a month on Render) or move to a proper database - flag this to me when you're ready and I'll help you upgrade it.
- **This is genuinely shared data.** Unlike the earlier click-through demo, everyone who logs in through this link sees and edits the same records - this is the real thing, not a preview.

## Adding your real team

Once the link is live:
1. Log in as the Owner (Vamshi Reddy, PIN `9999` - **change this PIN first**, see below).
2. Go to **People** → add each of your real supervisors, area supervisors, HR person, ops manager and data team members with their own PIN.
3. Go to **Assignments** → add your real sites, drivers, and vehicles, and assign each vehicle to a site, supervisor, driver and client.
4. Give each person their own name + PIN privately (don't share the same login between people - that's what makes the audit log meaningful).

**Change every demo PIN before real use.** The seeded accounts (Ravi, Suresh, Priya, Anil, Sneha, Vamshi, Divya) use simple 4-digit PINs for testing. Ask me to help you either edit them directly in `data/seed.json` before your first deploy, or add real replacement accounts through the People screen and deactivate the demo ones.

## If anything goes wrong

Send me a screenshot of the error (Render shows build/deploy logs on the service page), and I'll help you fix it from here.
