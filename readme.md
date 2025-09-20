# ProductivityTools.CalendarMonitor

This Google Apps Script project monitors a specified Google Calendar for meetings involving a particular person and a list of participants, then sends a summary email of today's and tomorrow's meetings.

## Features

- Monitors a calendar specified by email or calendar ID.
- Filters events to only those where exactly two people are invited: the main person and one participant from your list.
- Outputs matching events to a sheet named `MonitoredEvents`.
- Sends a daily summary email to `pwujczyk@google.com` with two sections: today's and tomorrow's meetings, showing only the meeting time in `HH:mm` format.

## Configuration

1. **Configuration Sheet Setup:**
   - Cell `A2`: Enter the email address or calendar ID of the person whose calendar you want to monitor.
   - Range `A7:A20`: Enter the email addresses of participants to monitor.

2. **Calendar Access:**
   - The calendar must be shared with your Google account (the one running the script) with at least "See all event details" permission.
   - Use the exact calendar ID (found in Google Calendar settings under "Integrate calendar") in cell `A2`.

## How It Works

- The script scans the specified calendar for events happening today and tomorrow.
- It checks if each event has exactly two guests and if at least one is from your participant list.
- Matching events are listed in the `MonitoredEvents` sheet and included in the summary email.

## Usage

1. Open the Google Apps Script editor attached to your Google Sheet.
2. Copy the code from `Code.js` into the script editor.
3. Set up your `Configuration` sheet as described above.
4. Run the `monitorCalendarAndDisplayEvents` function.

## Email Format

The summary email contains two sections:

- **Today Meetings**
- **Tomorrow Meetings**

Each section lists the meeting title and time in `HH:mm` format.

---

**Example Email:**

```
Today Meetings
Project Sync  09:00-09:30

Tomorrow Meetings
Design Review  14:00-14:30
```

---

## License

MIT License
