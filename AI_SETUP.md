# Turning on the AI Assistant

The app now has an **AI Assistant** (chat tab, Owner-only) and **AI Insights** (auto-generated cards on the Owner Dashboard). Both are powered by Anthropic's Claude API, which is a separate service from Render/MongoDB - it needs its own API key. Until you add one, both features show a clear "not set up yet" message instead of breaking anything else in the app.

This takes about 5 minutes.

## Part 1 - Get an API key

1. Go to **console.anthropic.com** and sign up (or log in if you already have an account).
2. You'll need to add billing details - the AI Assistant is pay-as-you-go, billed directly by Anthropic based on usage (see "What this costs" below). There's no separate subscription.
3. Go to **API Keys** (left sidebar) → **Create Key**. Give it a name like `padmasri-fleet-app` so you recognize it later.
4. Copy the key - it starts with `sk-ant-`. Save it somewhere safe; Anthropic only shows it to you once.

## Part 2 - Add it to Render

1. Log into **render.com** and open your `padmasri-travels` web service.
2. Go to **Environment** (left sidebar) → **Add Environment Variable**.
3. Add:
   - Key: `ANTHROPIC_API_KEY`
   - Value: the key you copied in Part 1 (starts with `sk-ant-`)
4. Click **Save Changes**. Render will automatically redeploy the app (takes 1-2 minutes).

That's it - once it redeploys, log in as Owner and the **✨ AI Assistant** tab and the **✨ AI Insights** card on the Dashboard will both start working.

## Testing it locally on your own computer

If you're running the app locally (not on Render), add the same key before starting the server:

```bash
export ANTHROPIC_API_KEY=sk-ant-your-key-here
node server.js
```

## Optional: changing the AI model

By default the app uses Anthropic's `claude-sonnet-4-20250514` model. If Anthropic releases a newer model later and you want to switch to it, you don't need new code - just add a second environment variable the same way as above:

- Key: `ANTHROPIC_MODEL`
- Value: the new model's name (Anthropic will publish this on their website)

## What this costs

Anthropic bills per question asked, based on how much data the AI reads to answer it - typically a fraction of a cent to a few cents per question for this app's scale of data. There's no fixed monthly fee. You can set a spending limit in the Anthropic console (**Settings → Limits**) if you want a hard cap. The AI Insights card is cached for an hour per refresh, so it doesn't re-run on every single Dashboard visit.

## What data the AI can see

The AI Assistant can only read fleet data (vehicles, costs, mileage, documents, drivers) through a fixed set of lookups - it cannot edit anything, cannot see passwords, and is only available to the Owner login. Anthropic's standard API terms apply to anything sent to it (the fleet data snapshots and your questions); Anthropic does not use API data to train its models by default.
