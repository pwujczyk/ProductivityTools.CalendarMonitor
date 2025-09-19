function monitorCalendarAndDisplayEvents() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configuration');
    var personEmail = sheet.getRange('A2').getValue();
    var participants = sheet.getRange('A7:A20').getValues().flat().filter(String);

    // Get calendar by email
    var calendar = CalendarApp.getCalendarById(personEmail);
    if (!calendar) {
        SpreadsheetApp.getUi().alert('Calendar not found for: ' + personEmail);
        return;
    }

    var now = new Date();
    var endOfTomorrow = new Date();
    endOfTomorrow.setDate(now.getDate() + 2);
    endOfTomorrow.setHours(0, 0, 0, 0);

    var events = calendar.getEvents(now, endOfTomorrow);
    var todayEvents = [];
    var tomorrowEvents = [];

    events.forEach(function(event) {
        var guests = event.getGuestList().map(function(guest) { return guest.getEmail(); });
        // Only consider events with exactly 2 people invited
        if (guests.length !== 2) return;
        var hasParticipant = participants.some(function(email) { return guests.includes(email); });
        if (hasParticipant) {
            var eventInfo = {
                title: event.getTitle(),
                start: event.getStartTime(),
                end: event.getEndTime(),
                guests: guests.join(', ')
            };
            var eventDate = event.getStartTime();
            var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
            var tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);

            if (eventDate >= today && eventDate < tomorrow) {
                todayEvents.push(eventInfo);
            } else if (eventDate >= tomorrow && eventDate < endOfTomorrow) {
                tomorrowEvents.push(eventInfo);
            }
        }
    });

    // Output results to a new sheet or clear and reuse an existing one
    var outputSheetName = 'MonitoredEvents';
    var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(outputSheetName);
    if (!outputSheet) {
        outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(outputSheetName);
    } else {
        outputSheet.clear();
    }

    outputSheet.appendRow(['Title', 'Start', 'End', 'Guests']);
    [].concat(todayEvents, tomorrowEvents).forEach(function(ev) {
        outputSheet.appendRow([ev.title, ev.start, ev.end, ev.guests]);
    });

    // Format time as HH:mm
    function formatTime(date) {
        var hours = date.getHours().toString().padStart(2, '0');
        var minutes = date.getMinutes().toString().padStart(2, '0');
        return hours + ':' + minutes;
    }

    // Prepare email content
    function formatEvents(events) {
        if (events.length === 0) return 'No meetings.<br>';
        return events.map(function(ev) {
            return '<b>' + ev.title + '</b><br>' +
                'Start: ' + formatTime(ev.start) + '<br>' +
                'End: ' + formatTime(ev.end) + '<br>' +
                'Guests: ' + ev.guests + '<br><br>';
        }).join('');
    }

    var emailBody =
        '<h2>Today Meetings</h2>' +
        formatEvents(todayEvents) +
        '<h2>Tomorrow Meetings</h2>' +
        formatEvents(tomorrowEvents);

    MailApp.sendEmail({
        to: 'pwujczyk@google.com',
        subject: 'Monitored Calendar Meetings: Today & Tomorrow',
        htmlBody: emailBody
    });
}
