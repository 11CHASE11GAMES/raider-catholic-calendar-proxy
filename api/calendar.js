const ICS_URL =
  "https://outlook.office365.com/owa/calendar/3f27e5fcd8c54156a67a04e6c92a556d@msoe.edu/39fd891e541a4016a9fecf8ed36628826223923538709763827/calendar.ics";

function unfoldICSLines(icsText) {
  return icsText.replace(/\r?\n[ \t]/g, "");
}

function unescapeICSValue(value) {
  return value
    .replace(/\\n/g, "\n")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\\\\/g, "\\");
}

function parseICSDate(rawValue) {
  if (!rawValue) {
    return null;
  }

  const value = rawValue.trim();

  if (/^\d{8}$/.test(value)) {
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(4, 6)) - 1;
    const day = Number(value.slice(6, 8));
    return new Date(year, month, day);
  }

  if (/^\d{8}T\d{6}Z$/.test(value)) {
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(4, 6)) - 1;
    const day = Number(value.slice(6, 8));
    const hour = Number(value.slice(9, 11));
    const minute = Number(value.slice(11, 13));
    const second = Number(value.slice(13, 15));
    return new Date(Date.UTC(year, month, day, hour, minute, second));
  }

  if (/^\d{8}T\d{6}$/.test(value)) {
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(4, 6)) - 1;
    const day = Number(value.slice(6, 8));
    const hour = Number(value.slice(9, 11));
    const minute = Number(value.slice(11, 13));
    const second = Number(value.slice(13, 15));
    return new Date(year, month, day, hour, minute, second);
  }

  return null;
}

function parseICS(icsText) {
  const unfolded = unfoldICSLines(icsText);
  const eventBlocks = unfolded.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];

  return eventBlocks
    .map((block) => {
      const lines = block.split(/\r?\n/);
      const event = {};

      lines.forEach((line) => {
        const separatorIndex = line.indexOf(":");

        if (separatorIndex === -1) {
          return;
        }

        const keyPart = line.slice(0, separatorIndex);
        const valuePart = line.slice(separatorIndex + 1);
        const baseKey = keyPart.split(";")[0];

        if (baseKey === "SUMMARY") {
          event.summary = unescapeICSValue(valuePart);
        }

        if (baseKey === "DESCRIPTION") {
          event.description = unescapeICSValue(valuePart);
        }

        if (baseKey === "LOCATION") {
          event.location = unescapeICSValue(valuePart);
        }

        if (baseKey === "DTSTART") {
          event.start = parseICSDate(valuePart);
        }

        if (baseKey === "DTEND") {
          event.end = parseICSDate(valuePart);
        }
      });

      return event;
    })
    .filter((event) => event.summary && event.start instanceof Date && !Number.isNaN(event.start.getTime()));
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

    const response = await fetch(ICS_URL, {
      method: "GET",
      headers: {
        "User-Agent": "Raider-Catholic-Calendar-Proxy"
      }
    });

    if (!response.ok) {
      return res.status(502).json({
        error: "Failed to fetch ICS feed",
        status: response.status
      });
    }

    const icsText = await response.text();
    const now = new Date();

    const upcomingEvents = parseICS(icsText)
      .filter((event) => {
        if (event.end instanceof Date && !Number.isNaN(event.end.getTime())) {
          return event.end >= now;
        }

        return event.start >= now;
      })
      .sort((a, b) => a.start - b.start)
      .slice(0, limit)
      .map((event) => ({
        summary: event.summary,
        description: event.description || "",
        location: event.location || "",
        start: event.start.toISOString(),
        end: event.end instanceof Date && !Number.isNaN(event.end.getTime())
          ? event.end.toISOString()
          : null
      }));

    return res.status(200).json({
      source: "outlook-ics",
      count: upcomingEvents.length,
      events: upcomingEvents
    });
  } catch (error) {
    return res.status(500).json({
      error: "Calendar proxy failed",
      message: error instanceof Error ? error.message : "Unknown error"
    });
  }
}