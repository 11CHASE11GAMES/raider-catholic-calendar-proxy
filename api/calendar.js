import ical from "node-ical";

const ICS_URL =
  "https://outlook.office365.com/owa/calendar/3f27e5fcd8c54156a67a04e6c92a556d@msoe.edu/39fd891e541a4016a9fecf8ed36628826223923538709763827/calendar.ics";

const LOOKAHEAD_DAYS = 90;

function stripHtml(value) {
  return value
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .trim();
}

function normalizeText(value) {
  if (!value) {
    return "";
  }

  return stripHtml(String(value))
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

function expandEvents(calendarData, fromDate, toDate) {
  const results = [];

  for (const entry of Object.values(calendarData)) {
    if (!entry || entry.type !== "VEVENT" || !(entry.start instanceof Date)) {
      continue;
    }

    if (entry.rrule) {
      const instances = ical.expandRecurringEvent(entry, {
        from: fromDate,
        to: toDate,
        includeOverrides: true,
        excludeExdates: true,
        expandOngoing: true
      });

      for (const instance of instances) {
        if (!(instance.start instanceof Date)) {
          continue;
        }

        const endDate =
          instance.end instanceof Date
            ? instance.end
            : entry.end instanceof Date
              ? new Date(instance.start.getTime() + (entry.end.getTime() - entry.start.getTime()))
              : null;

        results.push({
          summary: instance.summary || entry.summary || "Untitled Event",
          description: instance.description || entry.description || "",
          location: instance.location || entry.location || "",
          start: instance.start,
          end: endDate
        });
      }

      continue;
    }

    const effectiveEnd = entry.end instanceof Date ? entry.end : entry.start;

    if (effectiveEnd >= fromDate && entry.start <= toDate) {
      results.push({
        summary: entry.summary || "Untitled Event",
        description: entry.description || "",
        location: entry.location || "",
        start: entry.start,
        end: entry.end instanceof Date ? entry.end : null
      });
    }
  }

  return results;
}

export default async function handler(req, res) {
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
    const limit = Number.isFinite(limitParam) && limitParam > 0
      ? Math.min(limitParam, 20)
      : 5;

    const fromDate = new Date();
    const toDate = new Date(fromDate.getTime() + LOOKAHEAD_DAYS * 24 * 60 * 60 * 1000);

    const calendarData = await ical.async.fromURL(ICS_URL, {
      headers: {
        "User-Agent": "Raider-Catholic-Calendar-Proxy"
      }
    });

    const upcomingEvents = expandEvents(calendarData, fromDate, toDate)
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
      message: error instanceof Error ? error.message : "Unknown error"
    });
  }
}