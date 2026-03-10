const ical = require("node-ical");

const ICS_URL =
  "https://outlook.office365.com/owa/calendar/3f27e5fcd8c54156a67a04e6c92a556d@msoe.edu/39fd891e541a4016a9fecf8ed36628826223923538709763827/calendar.ics";

const LOOKAHEAD_DAYS = 90;

function normalizeText(value) {
  if (!value) {
    return "";
  }

  return String(value)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function serializeEvent(event) {
  return {
    summary: event.summary || "Untitled Event",
    description: normalizeText(event.description || ""),
    location: normalizeText(event.location || ""),
    start: event.start instanceof Date ? event.start.toISOString() : null,
    end: event.end instanceof Date ? event.end.toISOString() : null
  };
}

function fromURL(url, options = {}) {
  return new Promise((resolve, reject) => {
    ical.fromURL(url, options, (error, data) => {
      if (error) {
        reject(error);
        return;
      }

      resolve(data);
    });
  });
}

function buildRecurringInstances(event, fromDate, toDate) {
  const instances = [];

  if (!event.rrule || !(event.start instanceof Date)) {
    return instances;
  }

  const durationMs =
    event.end instanceof Date
      ? event.end.getTime() - event.start.getTime()
      : 0;

  const occurrenceDates = event.rrule.between(fromDate, toDate, true);

  occurrenceDates.forEach((occurrenceStart) => {
    const recurrenceId = occurrenceStart.toISOString().slice(0, 10);

    if (event.exdate) {
      const exdateKeys = Object.keys(event.exdate);
      const isExcluded = exdateKeys.some((key) => key.startsWith(recurrenceId));

      if (isExcluded) {
        return;
      }
    }

    instances.push({
      summary: event.summary || "Untitled Event",
      description: event.description || "",
      location: event.location || "",
      start: occurrenceStart,
      end: durationMs > 0 ? new Date(occurrenceStart.getTime() + durationMs) : null
    });
  });

  return instances;
}

function collectUpcomingEvents(calendarData, fromDate, toDate) {
  const events = [];

  Object.values(calendarData).forEach((entry) => {
    if (!entry || entry.type !== "VEVENT" || !(entry.start instanceof Date)) {
      return;
    }

    if (entry.rrule) {
      events.push(...buildRecurringInstances(entry, fromDate, toDate));
      return;
    }

    const effectiveEnd = entry.end instanceof Date ? entry.end : entry.start;

    if (effectiveEnd >= fromDate && entry.start <= toDate) {
      events.push({
        summary: entry.summary || "Untitled Event",
        description: entry.description || "",
        location: entry.location || "",
        start: entry.start,
        end: entry.end instanceof Date ? entry.end : null
      });
    }
  });

  return events;
}

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({
      error: "Method not allowed"
    });
  }

  try {
    const limitParam = Number(req.query.limit);
    const limit =
      Number.isFinite(limitParam) && limitParam > 0
        ? Math.min(limitParam, 20)
        : 5;

    const fromDate = new Date();
    const toDate = new Date(
      fromDate.getTime() + LOOKAHEAD_DAYS * 24 * 60 * 60 * 1000
    );

    const calendarData = await fromURL(ICS_URL, {
      headers: {
        "User-Agent": "Raider-Catholic-Calendar-Proxy"
      }
    });

    const upcomingEvents = collectUpcomingEvents(
      calendarData,
      fromDate,
      toDate
    )
      .filter((event) => event.start instanceof Date)
      .sort((a, b) => a.start - b.start)
      .slice(0, limit)
      .map(serializeEvent);

    return res.status(200).json({
      source: "outlook-ics",
      count: upcomingEvents.length,
      lookaheadDays: LOOKAHEAD_DAYS,
      events: upcomingEvents
    });
  } catch (error) {
    return res.status(500).json({
      error: "Calendar proxy failed",
      message: error && error.message ? error.message : "Unknown error"
    });
  }
};