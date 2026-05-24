
# ST-2110 Schedule Manager - Fast Search Version

Runs on port **5005**.

This version fixes the slow search problem by building an in-memory cache once at startup.
After the cache is ready, searching should be almost instant.

## Start

```bash
cd st2110_schedule_manager_fast_search
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Open:

```text
http://127.0.0.1:5005
```

## Important

On first startup the page will show:

```text
Building Excel cache...
```

After that, searches do not reopen the Excel file every time.

## Files

Workbook:

```text
data/Move_To_IP_Schedule_v5.96.xlsx
```

Backups:

```text
backups/
```


## Fixed Add Entry Version

Keeps the detailed view exactly as before and adds an **Add Entry** button. Select a specific sheet first, then click Add Entry.


## Delete Entry

Each record now has a **Delete** button. The app creates a backup before deleting a row from the Excel workbook.


## Working Unit Map

Clean self-contained Unit Map added. It does not depend on older map functions. Node names use Unit only, and drag/zoom/pan are implemented with pointer events.


## HTML Clickable Nodes

The Unit Map now uses normal HTML overlay nodes for source/destination circles. This avoids SVG click-target issues. Clicking a node shows details; clicking empty map clears panels; nodes remain draggable.


## Sort + Zoom Working Rebuild

Rebuilt from the confirmed working HTML-node map. Sort Map is added inside the working map code, and zoom/reset remain unchanged.


## Button Click Fix

Zoom, Reset, and Sort Map controls now stop event propagation and the map canvas ignores clicks on the control bar.


# V2 Operations Upgrade

Added practical first implementation of the recommended features:

1. IP conflict and duplicate detection
2. ST-2110 flow type inference: Video, Audio, ANC, PTP, NMOS, Control, Unknown
3. VLAN filtering and summary
4. Device database with ping status checks
5. NMOS configuration placeholder / preview endpoint
6. Audit log for key app actions
7. Export filtered XLSX with flow type and 2022-7 columns
8. Operations panel for validation and summaries
9. Devices panel for adding/checking devices
10. Existing Unit Map, import missing, export map, add/edit/delete retained


# Authentication and Roles

Added local portal authentication.

Default first login:
- Username: `admin`
- Password: `admin123`

Roles:
- `read`: view/search/map/export only
- `write`: add/edit/delete/import/device/status operations
- `admin`: all write permissions plus user management

Admin section:
- Create users
- Delete users
- Change user roles
- Force temporary password change on new users

Important:
Change the default admin password immediately after first login.
For production, set an environment variable:
`ST2110_SECRET_KEY=your-long-random-secret`


## Auth Import Fix

Fixed missing `import json` and checked required auth/operations imports.
