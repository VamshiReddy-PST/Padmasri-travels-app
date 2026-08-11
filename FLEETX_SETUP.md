# Turning on Live Tracking (Fleetx)

Every vehicle's own detail screen (My Vehicles, Reports → Cost Analysis, and the Owner's "Today's Status" pages) now has a **📍 Live Tracking** card - status, speed, location, and fuel level, pulled from your Fleetx GPS/fuel-sensor devices. This is a separate service from Render/MongoDB, so it needs its own token. Until it's set, the card just shows a plain "not set up yet" message instead of breaking anything else.

Vehicles are matched to Fleetx automatically by registration number - no per-vehicle setup needed. A vehicle that doesn't have a Fleetx GPS device on it yet simply shows "This vehicle isn't linked to a GPS device yet."

Camera feeds are **not** wired up yet - Fleetx currently only gives this app a camera device ID, not an actual video stream URL. The card will note when a vehicle has an onboard camera, but there's nothing to click yet. That's a follow-up once Fleetx/LOTIM's stream access details are available.

This takes about 2 minutes - you already have the token.

## Add it to Render

1. Log into **render.com** and open your `padmasri-travels` web service.
2. Go to **Environment** (left sidebar) → **Add Environment Variable**.
3. Add:
   - Key: `FLEETX_API_TOKEN`
   - Value: your Fleetx bearer token (the same one from the curl command you tested with)
4. Click **Save Changes**. Render will automatically redeploy the app (takes 1-2 minutes).

That's it - once it redeploys, every vehicle detail screen will start showing live tracking.

## Testing it locally on your own computer

If you're running the app locally (not on Render), add the same token before starting the server:

```bash
export FLEETX_API_TOKEN=your-fleetx-bearer-token
node server.js
```

## Optional: changing the Fleetx tag

By default the app asks Fleetx for the `Enmovil` tag, which is what returned your whole fleet in testing. If Fleetx ever reorganizes your account under a different tag name, you can override it without new code:

- Key: `FLEETX_TAG`
- Value: the new tag name

## A note on the token you already shared

Since the Fleetx bearer token was pasted directly into a chat conversation to test this integration, it's worth rotating it in Fleetx's dashboard once you've set the new one in Render - treat it the same as any password that was typed somewhere it didn't strictly need to be.
