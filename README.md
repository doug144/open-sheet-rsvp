# 💌 Dynamic Multi-Event RSVP Framework

An open-source, serverless, multi-event RSVP management system powered by **Google Sheets**, **Google Apps Script**, and a lightweight static **HTML/JS Frontend** (easily hosted on Cloudflare Pages, GitHub Pages, or Netlify).

Designed for weddings, conferences, family reunions, or weekend retreats where different guests are invited to different sub-events and respond as individuals or household groups.

---

## ✨ Key Features

- **Multi-Event Access Control:** Assign different events to different guests (e.g., Rehearsal Dinner for VIPs only, Main Reception for all guests).
- **Group & Household Bundling:** Unique URL parameters (`?id=GRP-101`) load all members of a family/group onto a single page so one person can RSVP for the entire party.
- **Dynamic "Plus One" Support:** Automatically handle +1 guests and allow recipients to submit their guest's full name.
- **Google Sheets Backend:** Manage your guest list, events, and configurations inside Google Sheets.
- **Automated Email Engine:**
  - Send custom HTML invitations with personalized RSVP links directly from Apps Script.
  - Send instant confirmation emails to guests containing one-click **Add to Google Calendar** links.
  - Send real-time submission alerts to event admins.
- **Batch Selection Shortcut:** One-click button to copy guest choices across all sub-events.
- **Custom Google Sheet Functions:** Includes custom spreadsheet formulas (like `=EVENT_SUMMARY()`) to track real-time attendance rollups.

---

## 🛠 Architecture Overview
```text
┌──────────────────────┐          GET / POST          ┌──────────────────────────┐
│                      │ ───────────────────────────> │                          │
│  Static Frontend     │                              │ Google Apps Script API   │
│ (Cloudflare Pages /  │ <─────────────────────────── │   (Web App Backend)      │
│   GitHub Pages)      │       JSON Responses         │                          │
└──────────────────────┘                              └─────────────┬────────────┘
│
Reads / Writes
│
▼
┌──────────────────────────┐
│   Google Sheets Database │
│ (Config, Events, Guests) │
└──────────────────────────┘
```

## 🚀 Quick Start & Setup Guide

### Step 1: Set Up the Google Sheet

1. Open the [Sample Google Sheet Template](https://docs.google.com/spreadsheets/d/1555WAipLnpX2-6KIOcPFmV24Jx4PDqgs5eluRh0gpOU).
2. Click **File ➔ Make a copy** to save it to your Google Drive.
3. Your sheet contains three primary tabs:
   - **`Config`**: Global settings (`event_title`, `event_dates`, `host_message`, `email_salutation`, `email_signature`, etc.).
   - **`Events`**: List of sub-events (`Event_ID`, `Title`, `Date_Time`, `Location`, `Description`, `Duration_Minutes`).
   - **`Guests`**: Guest database (`Guest_ID`, `Group_ID`, `Full_Name`, `Email`, `Allowed_Events`, `Plus_One_Allowed`, `RSVPs`, `Notes`, `Invite_Sent`).

---

### Step 2: Deploy the Backend (Google Apps Script)

1. Open your newly copied Google Sheet.
2. Go to **Extensions ➔ Apps Script**.
3. Replace any default code in `Code.gs` with the content from `/backend/Code.gs`.
4. **Configure Script Properties:**
   - Go to **Project Settings** (⚙️ gear icon on the left menu).
   - Scroll down to **Script Properties** and add:
     - `WEBSITE_URL`: The URL where your frontend site is hosted (e.g., `https://your-rsvp-site.pages.dev`).
     - `ADMIN_EMAIL`: Email address that will receive admin RSVP notifications.
     - `SEND_EMAILS`: Set to `true` to enable automated emails (`false` for testing).
5. **Deploy as Web App:**
   - Click **Deploy ➔ New deployment** top-right.
   - Select type: **Web app**.
   - Set **Execute as:** `Me`.
   - Set **Who has access:** `Anyone`.
   - Click **Deploy** and copy your **Web App URL** (looks like `https://script.google.com/macros/s/.../exec`).

---

### Step 3: Configure and Host the Frontend

1. Clone or download this repository.
1. In the `/frontend/` folder, copy `config.js.template` (or create a file named `config.js`) and set your deployed Web App URL:

    ```javascript
    const CONFIG = {
    scriptUrl: "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec",
    coverImageUrl: "https://example.com/your-cover-photo.jpg" // Optional
    };
    ```
1. Deploy the Frontend:
    Upload the `/frontend/` directory to any static site hosting provider. Cloudflare Pages is a great free option.

    **Example: Deploying to Cloudflare Pages**

    There are two common ways to deploy to Cloudflare Pages:

    *   **Direct Upload (Easiest):**
        1.  Log into your [Cloudflare dashboard](https://dash.cloudflare.com/).
        2.  Navigate to **Workers & Pages** ➔ **Create application** ➔ **Pages** ➔ **Upload assets**.
        3.  Give your project a name.
        4.  Drag and drop the entire `/frontend` folder from your computer into the upload box.
        5.  Click **Deploy site**.

    *   **Connect to Git (for automatic updates):**
        1.  Push your project code to a GitHub or GitLab repository.
        2.  In Cloudflare Pages, choose **Connect to Git** and select your repository.
        3.  In the build settings, set the **Build output directory** to `/frontend` and leave the **Build command** blank.
        4.  Click **Save and Deploy**.

## 📧 Sending Email Invitations
You can batch-send personalized HTML emails with dynamic links directly from Apps Script:

1. Populate your Guests tab with guest names, emails, assigned Group_IDs, and Allowed_Events.

1. Open Apps Script (Extensions ➔ Apps Script) from your google sheet.

1. Select the sendInvitations function in the editor toolbar and click Run.

1. The script will:

    - Generate a personalized link (https://your-site.com?id=GRP-101).

    - Send the invite via your GMail account.

    - Update the Invite_Sent column in your spreadsheet to Yes to prevent duplicate emails.

## 📊 Summary Spreadsheet Formula
The backend includes a custom formula function EVENT_SUMMARY() that computes live attendance counts.

In any empty cell in your Google Sheet, type:

```Excel
=EVENT_SUMMARY()
```

This auto-populates a real-time tracking table with Event Name, Invited, Attending, and Not Responded counts.

## 📁 Repository Structure
```text
├── backend/
│   └── Code.gs              # Apps Script backend API & email engine
├── frontend/
│   ├── index.html           # Main dynamic RSVP web application
│   ├── config.js            # Configuration file for Web App URL & cover image
│   └── config.js.template   # Template configuration file
└── README.md                # Documentation
```

## 📄 License
Distributed under the GNU General Public License v3.0 (GPL-3.0). See LICENSE for details.