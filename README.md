# ECM+ Outlook Add-in (self-hosted ASP.NET Core)

This project bundles everything in one place: the Outlook add-in's manifest,
task pane, and the API it talks to. Because they're served from the same
domain, you get:

- No CORS configuration needed for the add-in itself
- No mixed-content (http vs https) issues
- No ngrok / temporary URLs — one real domain, once

## Project layout

```
OutlookAddinProject/
  Program.cs                  <- serves wwwroot + API, in one process
  appsettings.json            <- your SQL connection string goes here
  Data/AppDbContext.cs        <- EF Core models (EmailRecord, AttachmentRecord)
  Controllers/EmailController.cs   <- /api/Emails/CreateDocumnet, /CreateCorrespondence
  wwwroot/
    manifest.xml              <- the add-in manifest (upload this to admin center)
    taskpane.html             <- the task pane UI (login, buttons, etc.)
    commands.html             <- required function-file placeholder
    assets/icon-*.png         <- ribbon icons
```

## 1. Before you run it

1. Open `appsettings.json` and set `ConnectionStrings:DefaultConnection` to your
   real SQL Server.
2. This project's `EmailController` is a minimal starting point (matches the
   task pane's existing calls). If you already have a fuller `EmailsController`
   /`AuthenticationController` in your real EcmPlus project (with AD login,
   MediatR, etc.), copy **those** controllers into this project's
   `Controllers/` folder instead, so you don't lose that logic. This scaffold
   is meant to show the *pattern* (one project, one domain) — merge it into
   your real backend rather than running both side by side.
3. Run EF Core migrations:
   ```
   dotnet ef migrations add InitialCreate
   dotnet ef database update
   ```

## 2. Publish it somewhere with a real HTTPS domain

Pick one:

- **Azure App Service** (easiest): `dotnet publish`, then deploy via Visual
  Studio's Publish wizard or `az webapp up`. You'll get a free
  `https://yourapp.azurewebsites.net` domain automatically.
- **IIS on your own server**: publish normally, and make sure the site has a
  valid SSL certificate bound to it (Let's Encrypt / win-acme works well).

Whichever you choose, note the final HTTPS domain — you'll need it next.

## 3. Point the manifest at your real domain

Open `wwwroot/manifest.xml` and replace **every** occurrence of
`YOUR-DOMAIN.com` with your actual published domain, e.g.:

```
https://myecmapp.azurewebsites.net
```

There are 6 places this appears: `IconUrl`, `HighResolutionIconUrl`,
`AppDomain`, `SourceLocation`, and the three `bt:Url`/`bt:Image` resources.

Re-publish the project after this change so the live `manifest.xml` reflects
the fix.

## 4. Upload the add-in to Outlook

### Option A — Deploy org-wide (Microsoft 365 admin center)

1. Go to **admin.cloud.microsoft** → **Settings** → **Integrated apps**.
2. Select **Upload custom apps**.
3. Choose **Office Add-in** as the app type, then **Upload manifest file**.
4. Browse to your published `manifest.xml` — or, since it's now hosted
   online, use **provide a URL** and paste
   `https://your-domain.com/manifest.xml` directly.
5. Choose who gets it (**Just me** for testing, or a group/everyone once
   confirmed working) and finish the deployment wizard.
6. It can take anywhere from a few minutes up to ~24 hours to appear in
   users' ribbons.

### Option B — Sideload to just your own account (instant, for testing)

1. In Outlook on the web, go to
   `outlook.office.com/mail/options/manageapps` (or the grid/apps icon in
   the ribbon → **My add-ins**).
2. Choose **Add a custom add-in** → **Add from file** (or **Add from URL**
   with your hosted `manifest.xml` link).
3. It appears immediately in your ribbon — no waiting.

## 5. Updating later

Any time you change `manifest.xml` (icons, buttons, URLs), remember to bump
the `<Version>` number — Outlook/admin center rejects a re-upload whose
version isn't strictly higher than what's already deployed.
