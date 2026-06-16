var DEFAULT_CONFIG = {
  displayName: "Prof. Hyunwoo Kim",
  calendarId: "primary",
  timeZone: "Asia/Seoul",
  showEventTitles: false,
  llm: {
    enabled: false,
    provider: "mindlogic",
    maxEvents: 3
  },
  mindlogic: {
    baseUrl: "https://factchat-cloud.mindlogic.ai/v1/gateway",
    model: "claude-sonnet-4-6"
  },
  openai: {
    enabled: false,
    model: "gpt-5.5"
  },
  officeHours: {
    mon: [["09:30", "12:00"], ["13:00", "17:00"]],
    tue: [["09:30", "12:00"], ["13:00", "17:00"]],
    wed: [["09:30", "12:00"], ["13:00", "17:00"]],
    thu: [["09:30", "12:00"], ["13:00", "17:00"]],
    fri: [["09:30", "12:00"], ["13:00", "17:00"]],
    sat: [],
    sun: []
  },
  focusTime: {
    enabled: true,
    start: "10:00",
    end: "12:00"
  },
  lunchBreak: {
    enabled: true,
    start: "12:00",
    end: "13:00"
  },
  officeEventKeywords: ["office hour", "office hours", "오피스아워", "상담"],
  officeLocationKeywords: ["경기도 일산동구 동국로 32"],
  awayKeywords: ["ooo", "out of office", "부재", "휴가", "출장"],
  titleStatusKeywords: {
    travel: ["출장", "외근", "business trip", "offsite"],
    leave: ["휴가", "연차", "반차", "vacation", "pto", "leave", "day off"],
    focus: ["focus", "집중", "방해 금지", "do not disturb"],
    meeting: ["회의", "미팅", "meeting", "mtg", "seminar", "세미나", "call"]
  }
};

var DAY_KEYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];

function refreshSenseCraftJson() {
  var config = getConfig_();
  var status = buildStatus_(config);
  var senseCraftJson = JSON.stringify(toSenseCraftStatus_(status, config), null, 2);
  publishToGitHub_(config, senseCraftJson);
  return senseCraftJson;
}

function installQuarterHourlyTrigger() {
  installRefreshTrigger_(15);
}

function installHourlyTrigger() {
  installQuarterHourlyTrigger();
}

function installRefreshTrigger_(minutes) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "refreshSenseCraftJson") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("refreshSenseCraftJson")
    .timeBased()
    .everyMinutes(minutes)
    .create();
}

function previewSenseCraftJson() {
  var config = getConfig_();
  var status = buildStatus_(config);
  Logger.log(JSON.stringify(toSenseCraftStatus_(status, config), null, 2));
}

function doPost(e) {
  var config = getConfig_();
  var body = parseRequestBody_(e);

  if (!config.location.webhookSecret) {
    return jsonResponse_({
      ok: false,
      error: "LOCATION_WEBHOOK_SECRET is not configured."
    });
  }

  if (String(body.secret || "") !== config.location.webhookSecret) {
    return jsonResponse_({
      ok: false,
      error: "Unauthorized."
    });
  }

  var presence = String(body.presence || body.status || "").toLowerCase();
  var distanceKm = null;

  if (body.latitude !== undefined && body.longitude !== undefined) {
    var lat = Number(body.latitude);
    var lng = Number(body.longitude);
    if (!isFinite(lat) || !isFinite(lng)) {
      return jsonResponse_({ ok: false, error: "Invalid latitude or longitude." });
    }
    if (!isFinite(config.location.officeLat) || !isFinite(config.location.officeLng)) {
      return jsonResponse_({ ok: false, error: "OFFICE_LAT and OFFICE_LNG are required." });
    }
    distanceKm = distanceBetweenKm_(lat, lng, config.location.officeLat, config.location.officeLng);
    presence = distanceKm > config.location.radiusKm ? "away" : "office";
  }

  if (presence !== "away" && presence !== "office") {
    return jsonResponse_({
      ok: false,
      error: "Use presence/status 'office' or 'away', or send latitude/longitude."
    });
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    GEOFENCE_STATUS: presence,
    GEOFENCE_UPDATED_AT: new Date().toISOString(),
    GEOFENCE_DISTANCE_KM: distanceKm === null ? "" : String(Math.round(distanceKm * 10) / 10)
  });

  var shouldPublish = String(body.publish === undefined ? "true" : body.publish).toLowerCase() !== "false";
  var published = false;
  var publishError = "";
  if (shouldPublish) {
    try {
      refreshSenseCraftJson();
      published = true;
    } catch (error) {
      publishError = error.message;
    }
  }

  return jsonResponse_({
    ok: !publishError,
    presence: presence,
    distance_km: distanceKm === null ? "" : Math.round(distanceKm * 10) / 10,
    published: published,
    publish_error: publishError,
    updated_at: new Date().toISOString()
  });
}

function doGet() {
  return jsonResponse_({ ok: true, service: "sensecraft-office-hours" });
}

function buildStatus_(config) {
  var now = new Date();
  var events = fetchCalendarEvents_(config, now);
  var status = computeStatus_(config, now, normalizeEvents_(events), {
    demo: false,
    configured: true,
    authorized: true
  });
  status.batteryLevel = fetchBatteryLevel_(config);
  return status;
}

function fetchCalendarEvents_(config, now) {
  var timeMin = startOfLocalWeekMonday_(now);
  var timeMax = addDays_(startOfLocalDay_(now), 8);
  var calendarId = config.calendarId || "primary";
  var calendar = calendarId === "primary"
    ? CalendarApp.getDefaultCalendar()
    : CalendarApp.getCalendarById(calendarId);

  if (!calendar) {
    throw new Error("Calendar not found or not accessible: " + calendarId);
  }

  return calendar
    .getEvents(timeMin, timeMax)
    .slice(0, 50)
    .map(function(event) {
      return calendarAppEventToApiEvent_(config, event);
    });
}

function calendarAppEventToApiEvent_(config, event) {
  var start = safeCalendarCall_(event, "getStartTime", null);
  var end = safeCalendarCall_(event, "getEndTime", null);
  var allDay = Boolean(safeCalendarCall_(event, "isAllDayEvent", false));

  var item = {
    id: safeCalendarCall_(event, "getId", ""),
    summary: safeCalendarCall_(event, "getTitle", ""),
    location: safeCalendarCall_(event, "getLocation", ""),
    status: "confirmed",
    eventType: "default",
    transparency: calendarEventTransparency_(event)
  };

  if (allDay) {
    item.start = { date: formatDateOnly_(start, config) };
    item.end = { date: formatDateOnly_(calendarAllDayEnd_(event, start, end), config) };
  } else {
    item.start = { dateTime: start.toISOString() };
    item.end = { dateTime: end.toISOString() };
  }

  return item;
}

function safeCalendarCall_(event, methodName, fallback) {
  try {
    return typeof event[methodName] === "function" ? event[methodName]() : fallback;
  } catch (error) {
    return fallback;
  }
}

function calendarEventTransparency_(event) {
  var transparency = String(safeCalendarCall_(event, "getTransparency", "opaque")).toLowerCase();
  return transparency.indexOf("transparent") !== -1 ? "transparent" : "opaque";
}

function calendarAllDayEnd_(event, start, end) {
  var allDayEnd = safeCalendarCall_(event, "getAllDayEndDate", null);
  if (allDayEnd) {
    return allDayEnd;
  }
  if (end && start && end > start) {
    return end;
  }
  return addDays_(start, 1);
}

function formatDateOnly_(date, config) {
  return Utilities.formatDate(date, config.timeZone, "yyyy-MM-dd");
}

function normalizeEvents_(events) {
  return events
    .filter(function(event) {
      return event.status !== "cancelled";
    })
    .map(function(event) {
      var range = readEventRange_(event);
      return {
        id: event.id,
        summary: event.summary || "",
        location: event.location || "",
        eventType: event.eventType || "default",
        transparency: event.transparency || "opaque",
        start: range.start,
        end: range.end,
        allDay: range.allDay,
        workingLocationProperties: event.workingLocationProperties,
        outOfOfficeProperties: event.outOfOfficeProperties,
        focusTimeProperties: event.focusTimeProperties
      };
    })
    .filter(function(event) {
      return event.start && event.end;
    })
    .sort(function(a, b) {
      return a.start - b.start;
    });
}

function readEventRange_(event) {
  if (event.start && event.start.dateTime) {
    return {
      start: new Date(event.start.dateTime),
      end: new Date(event.end.dateTime),
      allDay: false
    };
  }

  if (event.start && event.start.date) {
    return {
      start: dateOnlyToLocalDate_(event.start.date),
      end: dateOnlyToLocalDate_(event.end.date),
      allDay: true
    };
  }

  return { start: null, end: null, allDay: false };
}

function dateOnlyToLocalDate_(value) {
  var parts = value.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function computeStatus_(config, now, events, connection) {
  var currentEvents = events.filter(function(event) {
    return includesTime_(event, now);
  });
  var officeInfo = getOfficeInfo_(config, now, events);
  var currentAway = firstMatching_(currentEvents, function(event) {
    return isAwayEvent_(config, event);
  });
  var currentOffsite = firstMatching_(currentEvents, function(event) {
    return isOffsiteLocationEvent_(config, event);
  });
  var currentLlmStatus = null;
  var currentTitleStatus = firstTitleStatus_(config, currentEvents);
  var currentWorkingLocation = firstMatching_(currentEvents, isWorkingLocation_);
  var currentWorkingLocationInfo = currentWorkingLocation ? describeWorkingLocation_(currentWorkingLocation) : null;
  var currentFocus = firstMatching_(currentEvents, isFocusEvent_);
  var currentBusy = firstMatching_(currentEvents, function(event) {
    return isBusyEvent_(event) && !isOfficeEvent_(config, event);
  });
  var currentImplicitFocus = currentBusy ? null : implicitFocusTime_(config, now);
  var currentLunch = implicitLunchBreak_(config, now);

  var state = "closed";
  var headline = "Office Hours Closed";
  var detail = "Visits are not available right now.";
  var currentUntil = officeInfo.nextWindow ? officeInfo.nextWindow.start : null;

  if (!officeInfo.isOpen) {
    if (currentLunch && !currentAway && !currentOffsite && !currentTitleStatus &&
        !(currentWorkingLocationInfo && !currentWorkingLocationInfo.availableHere) && !currentBusy) {
      state = "lunch";
      headline = "Lunch Break";
      detail = "Back at " + formatTime_(currentLunch.end, config) + ".";
      currentUntil = currentLunch.end;
    } else {
      state = "off_hours";
      headline = "Out of Office";
      detail = officeInfo.nextWindow
        ? "Back " + formatSentenceDateTime_(officeInfo.nextWindow.start, now, config) + "."
        : "Office hours are closed.";
      currentUntil = officeInfo.nextWindow ? officeInfo.nextWindow.start : null;
    }
  } else if (isGeofenceAway_(config, now)) {
    state = "offsite";
    headline = "Out of Office";
    detail = "Away from the office.";
    currentUntil = null;
  } else if (currentAway) {
    state = "away";
    headline = "Away";
    detail = currentAway.allDay ? "Away for today." : "Away right now.";
    currentUntil = currentAway.end;
  } else if (currentOffsite) {
    state = "offsite";
    headline = "Out of Office";
    detail = "Offsite appointment.";
    currentUntil = currentOffsite.end;
  } else if ((currentLlmStatus = classifyCurrentStatusWithLlm_(config, now, currentEvents))) {
    state = currentLlmStatus.state;
    headline = currentLlmStatus.headline;
    detail = currentLlmStatus.detail;
    currentUntil = currentLlmStatus.currentUntil;
  } else if (currentTitleStatus) {
    state = currentTitleStatus.state;
    headline = currentTitleStatus.headline;
    detail = currentTitleStatus.detail;
    currentUntil = currentTitleStatus.event.end;
  } else if (currentWorkingLocationInfo && !currentWorkingLocationInfo.availableHere) {
    state = "remote";
    headline = "Working Remotely";
    detail = currentWorkingLocationInfo.label;
    currentUntil = currentWorkingLocation.end;
  } else if (currentFocus) {
    state = "focus";
    headline = "Focus Time";
    detail = "Please do not disturb.";
    currentUntil = currentFocus.end;
  } else if (currentImplicitFocus) {
    state = "focus";
    headline = "Focus Time";
    detail = "Please do not disturb.";
    currentUntil = currentImplicitFocus.end;
  } else if (currentBusy) {
    state = "busy";
    headline = "In a Meeting";
    detail = config.showEventTitles ? currentBusy.summary || "Busy" : "Busy right now.";
    currentUntil = currentBusy.end;
  } else if (currentWorkingLocationInfo) {
    state = "available";
    headline = "At the Office";
    detail = currentWorkingLocationInfo.label;
    currentUntil = currentWorkingLocation.end;
  } else if (officeInfo.isOpen) {
    state = "available";
    headline = "Available";
    detail = officeInfo.currentWindow && officeInfo.currentWindow.source === "calendar"
      ? "Office hours are open."
      : "Visits welcome.";
    currentUntil = officeInfo.currentWindow ? officeInfo.currentWindow.end : null;
  }

  var nextAvailableAt = findNextAvailable_(config, now, events);
  var agenda = events
    .filter(function(event) {
      return event.end > now;
    })
    .filter(function(event) {
      return !event.allDay;
    })
    .filter(function(event) {
      return event.start < addDays_(startOfLocalDay_(now), 1);
    })
    .filter(function(event) {
      return overlapsOfficeHours_(config, event);
    })
    .slice(0, 5)
    .map(function(event) {
      return presentAgendaEvent_(config, event);
    });

  return {
    connection: connection,
    displayName: config.displayName,
    generatedAt: new Date().toISOString(),
    now: now.toISOString(),
    state: state,
    headline: headline,
    detail: detail,
    currentUntil: currentUntil ? currentUntil.toISOString() : null,
    nextAvailableAt: nextAvailableAt ? nextAvailableAt.toISOString() : null,
    awayDays: weeklyAwayDays_(config, now, events),
    officeHoursText: officeHoursText_(config, now),
    agenda: agenda
  };
}

function firstMatching_(events, predicate) {
  return events.find(predicate) || null;
}

function includesTime_(event, date) {
  return event.start <= date && date < event.end;
}

function isAwayEvent_(config, event) {
  if (event.eventType === "outOfOffice") {
    return true;
  }
  return includesKeyword_(event.summary, config.awayKeywords || []);
}

function isFocusEvent_(event) {
  return event.eventType === "focusTime";
}

function isWorkingLocation_(event) {
  return event.eventType === "workingLocation";
}

function implicitFocusTime_(config, now) {
  var focusTime = config.focusTime || {};
  if (focusTime.enabled === false) {
    return null;
  }

  return implicitTimeWindow_(now, focusTime.start || "10:00", focusTime.end || "12:00");
}

function implicitLunchBreak_(config, now) {
  var lunchBreak = config.lunchBreak || {};
  if (lunchBreak.enabled === false) {
    return null;
  }

  return implicitTimeWindow_(now, lunchBreak.start || "12:00", lunchBreak.end || "13:00");
}

function implicitTimeWindow_(now, startValue, endValue) {
  var start = timeOnDate_(now, startValue);
  var end = timeOnDate_(now, endValue);
  if (!(start < end)) {
    return null;
  }
  return start <= now && now < end ? { start: start, end: end } : null;
}

function isOfficeEvent_(config, event) {
  return includesKeyword_(event.summary, config.officeEventKeywords || []);
}

function isOffsiteLocationEvent_(config, event) {
  var location = String(event.location || "").trim();
  return Boolean(location) && !isOfficeLocation_(config, location);
}

function weeklyAwayDays_(config, now, events) {
  var weekStart = startOfLocalWeekMonday_(now);
  var days = [];
  for (var offset = 0; offset < 5; offset += 1) {
    var day = addDays_(weekStart, offset);
    if (dayHasAwayBlocker_(config, day, events)) {
      days.push(DAY_KEYS[day.getDay()]);
    }
  }
  return days;
}

function dayHasAwayBlocker_(config, day, events) {
  var windows = officeWindowsForDate_(config, day);
  if (!windows.length) {
    return false;
  }

  return events.some(function(event) {
    if (!isAwayDayCandidate_(config, event)) {
      return false;
    }
    if (event.allDay) {
      return eventOverlapsDay_(event, day);
    }
    return windows.some(function(window) {
      return event.start < window.end && window.start < event.end;
    });
  });
}

function isAwayDayCandidate_(config, event) {
  if (isAwayEvent_(config, event) || isOffsiteLocationEvent_(config, event)) {
    return true;
  }

  var titleStatus = titleStatusForEvent_(config, event);
  if (titleStatus && ["away", "leave", "offsite"].indexOf(titleStatus.state) !== -1) {
    return true;
  }

  if (isWorkingLocation_(event)) {
    var workingLocation = describeWorkingLocation_(event);
    return !workingLocation.availableHere;
  }

  return false;
}

function eventOverlapsDay_(event, day) {
  var dayEnd = addDays_(day, 1);
  return event.start < dayEnd && day < event.end;
}

function firstTitleStatus_(config, events) {
  for (var index = 0; index < events.length; index += 1) {
    var status = titleStatusForEvent_(config, events[index]);
    if (status) {
      status.event = events[index];
      return status;
    }
  }
  return null;
}

function classifyCurrentStatusWithLlm_(config, now, currentEvents) {
  if (!config.llm.enabled) {
    return null;
  }

  var candidates = currentEvents
    .filter(function(event) {
      return !isOfficeEvent_(config, event) && !isWorkingLocation_(event);
    })
    .slice(0, config.llm.maxEvents);

  if (!candidates.length) {
    return null;
  }

  try {
    var input = {
      now: now.toISOString(),
      time_zone: config.timeZone,
      office_location_keywords: config.officeLocationKeywords || [],
      events: candidates.map(function(event) {
        return {
          id: event.id || "",
          title: event.summary || "",
          location: event.location || "",
          event_type: event.eventType || "default",
          transparency: event.transparency || "opaque",
          all_day: Boolean(event.allDay),
          start: event.start.toISOString(),
          end: event.end.toISOString()
        };
      })
    };

    var classification = requestLlmClassification_(config, input);
    if (!classification || !classification.should_override) {
      return null;
    }

    var status = statusFromLlmClassification_(classification);
    if (!status) {
      return null;
    }

    status.currentUntil = currentUntilForClassification_(classification, candidates);
    return status;
  } catch (error) {
    return null;
  }
}

function requestLlmClassification_(config, input) {
  if (config.llm.provider === "mindlogic") {
    return requestMindlogicClassification_(config, input);
  }
  if (config.llm.provider === "openai") {
    return requestOpenAiClassification_(config, input);
  }
  return null;
}

function requestMindlogicClassification_(config, input) {
  if (!config.mindlogic.apiKey) {
    return null;
  }

  var url = trimTrailingSlash_(config.mindlogic.baseUrl) + "/chat/completions/";
  var payload = {
    model: config.mindlogic.model,
    messages: [
      {
        role: "system",
        content: llmStatusInstructions_() + " Respond with only a compact JSON object matching this schema: " +
          JSON.stringify(llmStatusSchema_())
      },
      {
        role: "user",
        content: JSON.stringify(input)
      }
    ],
    max_tokens: 300,
    temperature: 0,
    response_format: {
      type: "json_object"
    }
  };

  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + config.mindlogic.apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    delete payload.response_format;
    response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + config.mindlogic.apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      return null;
    }
  }

  var body = JSON.parse(response.getContentText() || "{}");
  var content = body.choices &&
    body.choices[0] &&
    body.choices[0].message &&
    body.choices[0].message.content;
  return parseClassificationJson_(content);
}

function requestOpenAiClassification_(config, input) {
  if (!config.openai.apiKey) {
    return null;
  }

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + config.openai.apiKey
    },
    payload: JSON.stringify({
      model: config.openai.model,
      store: false,
      max_output_tokens: 300,
      instructions: llmStatusInstructions_(),
      input: JSON.stringify(input),
      text: {
        format: {
          type: "json_schema",
          name: "office_door_status",
          strict: true,
          schema: llmStatusSchema_()
        }
      }
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    return null;
  }

  var body = JSON.parse(response.getContentText() || "{}");
  return parseClassificationJson_(extractOpenAiOutputText_(body));
}

function parseClassificationJson_(value) {
  if (!value) {
    return null;
  }
  if (typeof value === "object") {
    return value;
  }

  var text = String(value).trim();
  try {
    return JSON.parse(text);
  } catch (error) {
    var start = text.indexOf("{");
    var end = text.lastIndexOf("}");
    if (start === -1 || end === -1 || end <= start) {
      return null;
    }
    try {
      return JSON.parse(text.slice(start, end + 1));
    } catch (nestedError) {
      return null;
    }
  }
}

function llmStatusInstructions_() {
  return [
    "Classify Prof. Hyunwoo Kim's current office-door status from calendar events.",
    "The display is public, so never reveal event titles, private names, or exact locations.",
    "Return only the structured JSON schema fields.",
    "Use Korean and English context in titles.",
    "Prefer the most restrictive accurate status.",
    "Use offsite for business trips, external appointments, conferences, seminars away from the office, or non-office locations.",
    "Use leave for vacation, PTO, annual leave, sick leave, personal leave, or all-day leave.",
    "Use focus for focus time or do-not-disturb work.",
    "Use busy for meetings, calls, interviews, advising, seminars, reviews, and opaque appointments.",
    "Use remote for working from home or remote work.",
    "Use away for generic out-of-office or unavailable status that is not leave or travel.",
    "If the events should not change the display status, set should_override to false and state to available."
  ].join(" ");
}

function llmStatusSchema_() {
  return {
    type: "object",
    additionalProperties: false,
    properties: {
      should_override: {
        type: "boolean",
        description: "Whether this classification should override normal availability."
      },
      state: {
        type: "string",
        enum: ["available", "busy", "away", "leave", "focus", "remote", "offsite"]
      },
      detail_kind: {
        type: "string",
        enum: ["available", "busy", "away", "leave", "focus", "remote", "offsite", "travel"]
      },
      event_id: {
        type: "string",
        description: "The event id that best explains the classification, or an empty string."
      }
    },
    required: ["should_override", "state", "detail_kind", "event_id"]
  };
}

function extractOpenAiOutputText_(body) {
  if (body.output_text) {
    return body.output_text;
  }

  var output = body.output || [];
  for (var outputIndex = 0; outputIndex < output.length; outputIndex += 1) {
    var content = output[outputIndex].content || [];
    for (var contentIndex = 0; contentIndex < content.length; contentIndex += 1) {
      var item = content[contentIndex];
      if ((item.type === "output_text" || item.type === "text") && item.text) {
        return item.text;
      }
    }
  }

  return "";
}

function trimTrailingSlash_(value) {
  return String(value || "").replace(/\/+$/, "");
}

function statusFromLlmClassification_(classification) {
  var state = String(classification.state || "");
  var detailKind = String(classification.detail_kind || state);

  var detailByKind = {
    available: {
      state: "available",
      headline: "Available",
      detail: "Visits welcome."
    },
    busy: {
      state: "busy",
      headline: "In a Meeting",
      detail: "Busy right now."
    },
    away: {
      state: "away",
      headline: "Away",
      detail: "Away right now."
    },
    leave: {
      state: "leave",
      headline: "On Leave",
      detail: "On leave right now."
    },
    focus: {
      state: "focus",
      headline: "Focus Time",
      detail: "Please do not disturb."
    },
    remote: {
      state: "remote",
      headline: "Working Remotely",
      detail: "Working remotely."
    },
    offsite: {
      state: "offsite",
      headline: "Out of Office",
      detail: "Offsite appointment."
    },
    travel: {
      state: "offsite",
      headline: "Out of Office",
      detail: "Business trip."
    }
  };

  var status = detailByKind[detailKind] || detailByKind[state];
  if (!status || status.state === "available") {
    return null;
  }

  return {
    state: status.state,
    headline: status.headline,
    detail: status.detail,
    currentUntil: null
  };
}

function currentUntilForClassification_(classification, events) {
  var eventId = String(classification.event_id || "");
  var matched = events.find(function(event) {
    return String(event.id || "") === eventId;
  });
  if (matched) {
    return matched.end;
  }

  return events.reduce(function(latest, event) {
    return !latest || event.end > latest ? event.end : latest;
  }, null);
}

function titleStatusForEvent_(config, event) {
  if (isOfficeEvent_(config, event)) {
    return null;
  }

  var title = event.summary || "";
  var keywords = config.titleStatusKeywords || {};
  if (includesKeyword_(title, keywords.travel || [])) {
    return {
      state: "offsite",
      headline: "Out of Office",
      detail: event.allDay ? "Business trip today." : "Business trip."
    };
  }
  if (includesKeyword_(title, keywords.leave || [])) {
    return {
      state: "leave",
      headline: "On Leave",
      detail: event.allDay ? "On leave today." : "On leave right now."
    };
  }
  if (includesKeyword_(title, keywords.focus || [])) {
    return {
      state: "focus",
      headline: "Focus Time",
      detail: "Please do not disturb."
    };
  }
  if (includesKeyword_(title, keywords.meeting || [])) {
    return {
      state: "busy",
      headline: "In a Meeting",
      detail: "Busy right now."
    };
  }
  return null;
}

function isTitleBlockingEvent_(config, event) {
  return Boolean(titleStatusForEvent_(config, event));
}

function isBusyEvent_(event) {
  if (event.eventType === "workingLocation") {
    return false;
  }
  if (event.eventType === "outOfOffice" || event.eventType === "focusTime") {
    return true;
  }
  return event.transparency !== "transparent";
}

function includesKeyword_(text, keywords) {
  var lower = String(text || "").toLowerCase();
  return keywords.some(function(keyword) {
    return lower.indexOf(String(keyword).toLowerCase()) !== -1;
  });
}

function isOfficeLocation_(config, location) {
  var officeLocations = config.officeLocationKeywords || [];
  if (!officeLocations.length) {
    return false;
  }

  var normalizedLocation = normalizeLocationText_(location);
  return officeLocations.some(function(officeLocation) {
    return normalizedLocation.indexOf(normalizeLocationText_(officeLocation)) !== -1;
  });
}

function normalizeLocationText_(value) {
  return String(value || "").toLowerCase().replace(/[\s,.-]/g, "");
}

function describeWorkingLocation_(event) {
  var props = event.workingLocationProperties || {};
  if (props.type === "officeLocation") {
    return {
      availableHere: true,
      label: props.officeLocation && props.officeLocation.label
        ? props.officeLocation.label
        : "At the office."
    };
  }
  if (props.type === "homeOffice") {
    return {
      availableHere: false,
      label: "Working from home."
    };
  }
  var label = props.customLocation && props.customLocation.label
    ? props.customLocation.label
    : event.location || "Working away from the office.";
  return {
    availableHere: includesKeyword_(label, ["office", "사무실", "연구실"]),
    label: label
  };
}

function getOfficeInfo_(config, now, events) {
  var calendarWindows = events
    .filter(function(event) {
      return isOfficeEvent_(config, event);
    })
    .map(function(event) {
      return {
        start: event.start,
        end: event.end,
        source: "calendar"
      };
    });

  var staticWindows = officeWindowsForDate_(config, now).map(function(window) {
    return {
      start: window.start,
      end: window.end,
      source: "schedule"
    };
  });

  var windows = calendarWindows.concat(staticWindows).sort(function(a, b) {
    return a.start - b.start;
  });
  var currentWindow = windows.find(function(window) {
    return window.start <= now && now < window.end;
  }) || null;
  var nextWindow = windows.find(function(window) {
    return window.start > now;
  }) || nextOfficeWindowAfter_(config, now);

  return {
    isOpen: Boolean(currentWindow),
    currentWindow: currentWindow,
    nextWindow: nextWindow
  };
}

function officeWindowsForDate_(config, date) {
  var key = DAY_KEYS[date.getDay()];
  var windows = (config.officeHours && config.officeHours[key]) || [];
  return windows.map(function(window) {
    return {
      start: timeOnDate_(date, window[0]),
      end: timeOnDate_(date, window[1])
    };
  });
}

function officeHoursText_(config, date) {
  var windows = officeWindowsForDate_(config, date);
  if (!windows.length) {
    return "No hours today";
  }
  return windows.map(function(window) {
    return formatTime_(window.start, config) + "-" + formatTime_(window.end, config);
  }).join(" / ");
}

function overlapsOfficeHours_(config, event) {
  var day = startOfLocalDay_(event.start);
  var lastDay = startOfLocalDay_(event.end);

  while (day <= lastDay) {
    if (officeWindowsForDate_(config, day).some(function(window) {
      return event.start < window.end && window.start < event.end;
    })) {
      return true;
    }
    day = addDays_(day, 1);
  }

  return false;
}

function nextOfficeWindowAfter_(config, date) {
  for (var offset = 0; offset < 14; offset += 1) {
    var day = addDays_(startOfLocalDay_(date), offset);
    var windows = officeWindowsForDate_(config, day);
    var next = windows.find(function(window) {
      return window.end > date;
    });
    if (next) {
      return {
        start: next.start,
        end: next.end,
        source: "schedule"
      };
    }
  }
  return null;
}

function timeOnDate_(date, hhmm) {
  var parts = hhmm.split(":").map(Number);
  var result = new Date(date);
  result.setHours(parts[0], parts[1], 0, 0);
  return result;
}

function findNextAvailable_(config, now, events) {
  var candidate = new Date(now);

  for (var attempts = 0; attempts < 40; attempts += 1) {
    var officeInfo = getOfficeInfo_(config, candidate, events);
    if (!officeInfo.isOpen) {
      if (!officeInfo.nextWindow) {
        return null;
      }
      candidate = new Date(officeInfo.nextWindow.start);
      continue;
    }

    var implicitBlocker = implicitAvailabilityBlocker_(config, candidate);
    if (implicitBlocker) {
      candidate = new Date(implicitBlocker.end);
      continue;
    }

    var blocker = events.find(function(event) {
      return includesTime_(event, candidate) &&
        ((isBusyEvent_(event) && !isOfficeEvent_(config, event)) ||
          isOffsiteLocationEvent_(config, event) ||
          isTitleBlockingEvent_(config, event));
    });
    if (blocker) {
      candidate = new Date(blocker.end);
      continue;
    }

    return candidate;
  }

  return null;
}

function implicitAvailabilityBlocker_(config, date) {
  return implicitFocusTime_(config, date) || implicitLunchBreak_(config, date);
}

function presentAgendaEvent_(config, event) {
  var type = agendaType_(config, event);
  var title = config.showEventTitles && event.summary ? event.summary : type.label;
  return {
    id: event.id,
    title: title,
    type: type.key,
    start: event.start.toISOString(),
    end: event.end.toISOString(),
    allDay: event.allDay
  };
}

function agendaType_(config, event) {
  if (isAwayEvent_(config, event)) {
    return { key: "away", label: "Away" };
  }
  if (isOfficeEvent_(config, event)) {
    return { key: "office", label: "Office Hours" };
  }
  if (isWorkingLocation_(event)) {
    return { key: "location", label: "Work Location" };
  }
  if (isFocusEvent_(event)) {
    return { key: "focus", label: "Focus Time" };
  }
  if (isBusyEvent_(event)) {
    return { key: "busy", label: "Busy" };
  }
  return { key: "free", label: "Available" };
}

function toSenseCraftStatus_(status, config) {
  var agenda = status.agenda || [];
  var first = agenda[0] || null;
  var second = agenda[1] || null;
  var third = agenda[2] || null;

  return {
    display_name: status.displayName || "",
    date: formatDateLabel_(status.now, config),
    state: status.state || "",
    status_label: senseCraftStateLabel_(status.state),
    headline: status.headline || "",
    detail: status.detail || "",
    next_available: formatRelativeDateTime_(status.nextAvailableAt, status.now, config),
    current_until: formatRelativeDateTime_(status.currentUntil, status.now, config),
    todays_hours: status.officeHoursText || "",
    away_days: status.awayDays || [],
    battery_level: formatBatteryLevel_(status.batteryLevel),
    up_next_1_time: formatAgendaRange_(first, config),
    up_next_1_title: first && first.title ? first.title : "",
    up_next_1_type: first && first.type ? first.type : "",
    up_next_2_time: formatAgendaRange_(second, config),
    up_next_2_title: second && second.title ? second.title : "",
    up_next_2_type: second && second.type ? second.type : "",
    up_next_3_time: formatAgendaRange_(third, config),
    up_next_3_title: third && third.title ? third.title : "",
    up_next_3_type: third && third.type ? third.type : "",
    updated_at: formatDateTimeLabel_(status.generatedAt, config),
    source: status.connection && status.connection.demo ? "demo" : "calendar"
  };
}

function fetchBatteryLevel_(config) {
  if (!config.seeed.url || !config.seeed.apiKey || !config.seeed.dataKey) {
    return null;
  }

  try {
    var response = UrlFetchApp.fetch(config.seeed.url, {
      method: "get",
      headers: {
        "api-key": config.seeed.apiKey
      },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      return null;
    }
    var body = JSON.parse(response.getContentText() || "{}");
    return readPath_(body, config.seeed.dataKey);
  } catch (error) {
    return null;
  }
}

function publishToGitHub_(config, content) {
  if (!config.github.token || !config.github.owner || !config.github.repo || !config.github.path) {
    throw new Error("GitHub Script Properties are incomplete.");
  }

  var apiPath = encodeURIComponent(config.github.path).replace(/%2F/g, "/");
  var url = "https://api.github.com/repos/" +
    encodeURIComponent(config.github.owner) +
    "/" +
    encodeURIComponent(config.github.repo) +
    "/contents/" +
    apiPath;

  var existing = UrlFetchApp.fetch(url + "?ref=" + encodeURIComponent(config.github.branch), {
    method: "get",
    headers: githubHeaders_(config),
    muteHttpExceptions: true
  });

  var sha = "";
  if (existing.getResponseCode() === 200) {
    sha = JSON.parse(existing.getContentText()).sha || "";
  } else if (existing.getResponseCode() !== 404) {
    throw new Error("Could not read GitHub file: " + existing.getContentText());
  }

  var payload = {
    message: "Update SenseCraft status JSON",
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: config.github.branch
  };
  if (sha) {
    payload.sha = sha;
  }

  var updated = UrlFetchApp.fetch(url, {
    method: "put",
    headers: githubHeaders_(config),
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (updated.getResponseCode() < 200 || updated.getResponseCode() >= 300) {
    throw new Error("Could not update GitHub file: " + updated.getContentText());
  }
}

function githubHeaders_(config) {
  return {
    Authorization: "Bearer " + config.github.token,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "sensecraft-office-hours-apps-script"
  };
}

function isGeofenceAway_(config, now) {
  var props = PropertiesService.getScriptProperties();
  var status = String(props.getProperty("GEOFENCE_STATUS") || "").toLowerCase();
  var updatedAtText = props.getProperty("GEOFENCE_UPDATED_AT");
  if (status !== "away" || !updatedAtText) {
    return false;
  }

  var updatedAt = new Date(updatedAtText);
  if (!isFinite(updatedAt.getTime())) {
    return false;
  }

  var ageMinutes = (now.getTime() - updatedAt.getTime()) / 60000;
  return ageMinutes <= config.location.staleMinutes;
}

function distanceBetweenKm_(lat1, lng1, lat2, lng2) {
  var earthKm = 6371;
  var dLat = degreesToRadians_(lat2 - lat1);
  var dLng = degreesToRadians_(lng2 - lng1);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(degreesToRadians_(lat1)) *
      Math.cos(degreesToRadians_(lat2)) *
      Math.sin(dLng / 2) *
      Math.sin(dLng / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return earthKm * c;
}

function degreesToRadians_(degrees) {
  return degrees * Math.PI / 180;
}

function readPath_(value, pathText) {
  return String(pathText || "")
    .split(".")
    .filter(Boolean)
    .reduce(function(current, key) {
      return current && current[key] !== undefined ? current[key] : null;
    }, value);
}

function formatBatteryLevel_(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  var numeric = Number(value);
  if (!isFinite(numeric)) {
    return String(value);
  }

  var percent = numeric <= 1 ? numeric * 100 : numeric;
  return String(Math.round(Math.max(0, Math.min(100, percent)))) + "%";
}

function senseCraftStateLabel_(state) {
  var labels = {
    available: "Available",
    busy: "Busy",
    away: "Away",
    leave: "On Leave",
    focus: "Do Not Disturb",
    lunch: "Lunch Break",
    remote: "Remote",
    off_hours: "Out of Office",
    offsite: "Out of Office",
    closed: "Closed",
    setup: "Setup Needed"
  };
  return labels[state] || "Status";
}

function formatDateLabel_(value, config) {
  if (!value) {
    return "";
  }
  return Utilities.formatDate(new Date(value), config.timeZone, "EEE, MMM d");
}

function formatDateTimeLabel_(value, config) {
  if (!value) {
    return "";
  }
  return Utilities.formatDate(new Date(value), config.timeZone, "MMM d, HH:mm");
}

function formatRelativeDateTime_(value, nowValue, config) {
  if (!value) {
    return "";
  }

  var date = new Date(value);
  var now = nowValue ? new Date(nowValue) : new Date();
  if (sameLocalDate_(date, now, config)) {
    return formatTime_(date, config);
  }

  var tomorrow = addDays_(now, 1);
  if (sameLocalDate_(date, tomorrow, config)) {
    return "Tomorrow " + formatTime_(date, config);
  }

  return formatDateTimeLabel_(value, config);
}

function formatSentenceDateTime_(value, nowValue, config) {
  return formatRelativeDateTime_(value, nowValue, config)
    .replace(/^Today\b/, "today")
    .replace(/^Tomorrow\b/, "tomorrow");
}

function formatAgendaRange_(item, config) {
  if (!item) {
    return "";
  }
  if (item.allDay) {
    return "All Day";
  }
  return formatTime_(new Date(item.start), config) + "-" + formatTime_(new Date(item.end), config);
}

function formatTime_(date, config) {
  return Utilities.formatDate(date, config.timeZone, "HH:mm");
}

function sameLocalDate_(a, b, config) {
  return Utilities.formatDate(a, config.timeZone, "yyyy-MM-dd") ===
    Utilities.formatDate(b, config.timeZone, "yyyy-MM-dd");
}

function addDays_(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function startOfLocalWeekMonday_(date) {
  var result = startOfLocalDay_(date);
  var daysSinceMonday = (result.getDay() + 6) % 7;
  result.setDate(result.getDate() - daysSinceMonday);
  return result;
}

function startOfLocalDay_(date) {
  var result = new Date(date);
  result.setHours(0, 0, 0, 0);
  return result;
}

function parseRequestBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }
  var type = e.postData.type || "";
  if (type.indexOf("application/json") !== -1) {
    return JSON.parse(e.postData.contents || "{}");
  }

  var output = {};
  e.postData.contents.split("&").forEach(function(pair) {
    var parts = pair.split("=");
    if (parts[0]) {
      output[decodeURIComponent(parts[0])] = decodeURIComponent(parts.slice(1).join("=") || "");
    }
  });
  return output;
}

function jsonResponse_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  var config = JSON.parse(JSON.stringify(DEFAULT_CONFIG));

  config.displayName = prop_(props, "DISPLAY_NAME", config.displayName);
  config.calendarId = prop_(props, "CALENDAR_ID", config.calendarId);
  config.timeZone = prop_(props, "TIME_ZONE", config.timeZone);
  config.showEventTitles = propBool_(props, "SHOW_EVENT_TITLES", config.showEventTitles);
  config.officeHours = propJson_(props, "OFFICE_HOURS_JSON", config.officeHours);
  config.focusTime = propJson_(props, "FOCUS_TIME_JSON", config.focusTime);
  config.lunchBreak = propJson_(props, "LUNCH_BREAK_JSON", config.lunchBreak);
  config.officeEventKeywords = propCsv_(props, "OFFICE_EVENT_KEYWORDS", config.officeEventKeywords);
  config.officeLocationKeywords = propCsv_(props, "OFFICE_LOCATION_KEYWORDS", config.officeLocationKeywords);
  config.awayKeywords = propCsv_(props, "AWAY_KEYWORDS", config.awayKeywords);
  config.titleStatusKeywords = {
    travel: propCsv_(props, "TITLE_TRAVEL_KEYWORDS", config.titleStatusKeywords.travel),
    leave: propCsv_(props, "TITLE_LEAVE_KEYWORDS", config.titleStatusKeywords.leave),
    focus: propCsv_(props, "TITLE_FOCUS_KEYWORDS", config.titleStatusKeywords.focus),
    meeting: propCsv_(props, "TITLE_MEETING_KEYWORDS", config.titleStatusKeywords.meeting)
  };

  var mindlogicApiKey = prop_(props, "MINDLOGIC_API_KEY", "");
  var openAiApiKey = prop_(props, "OPENAI_API_KEY", "");
  var defaultLlmProvider = mindlogicApiKey ? "mindlogic" : openAiApiKey ? "openai" : config.llm.provider;
  config.llm = {
    provider: String(prop_(props, "LLM_PROVIDER", defaultLlmProvider)).toLowerCase(),
    enabled: propBool_(props, "LLM_ENABLED", Boolean(mindlogicApiKey || openAiApiKey)),
    maxEvents: Number(prop_(props, "LLM_MAX_EVENTS", prop_(props, "OPENAI_MAX_EVENTS", "3")))
  };
  if (!isFinite(config.llm.maxEvents) || config.llm.maxEvents < 1) {
    config.llm.maxEvents = 3;
  }

  config.mindlogic = {
    apiKey: mindlogicApiKey,
    baseUrl: prop_(props, "MINDLOGIC_BASE_URL", config.mindlogic.baseUrl),
    model: prop_(props, "MINDLOGIC_MODEL", config.mindlogic.model)
  };

  config.openai = {
    apiKey: openAiApiKey,
    model: prop_(props, "OPENAI_MODEL", config.openai.model),
    enabled: propBool_(props, "OPENAI_ENABLED", Boolean(openAiApiKey))
  }

  config.github = {
    owner: prop_(props, "GITHUB_OWNER", "ChungmaruQ"),
    repo: prop_(props, "GITHUB_REPO", "desktop-tutorial"),
    branch: prop_(props, "GITHUB_BRANCH", "master"),
    path: prop_(props, "GITHUB_PATH", "sensecraft.json"),
    token: prop_(props, "GITHUB_TOKEN", "")
  };

  config.seeed = {
    url: prop_(props, "SEEED_BATTERY_URL", ""),
    dataKey: prop_(props, "SEEED_BATTERY_DATA_KEY", "result.battery.level"),
    apiKey: prop_(props, "SEEED_BATTERY_API_KEY", "")
  };

  config.location = {
    webhookSecret: prop_(props, "LOCATION_WEBHOOK_SECRET", ""),
    officeLat: Number(prop_(props, "OFFICE_LAT", "")),
    officeLng: Number(prop_(props, "OFFICE_LNG", "")),
    radiusKm: Number(prop_(props, "OFFICE_RADIUS_KM", "10")),
    staleMinutes: Number(prop_(props, "LOCATION_STALE_MINUTES", "1440"))
  };

  return config;
}

function prop_(props, key, fallback) {
  var value = props.getProperty(key);
  return value === null || value === undefined || value === "" ? fallback : value;
}

function propBool_(props, key, fallback) {
  var value = props.getProperty(key);
  if (value === null || value === undefined || value === "") {
    return fallback;
  }
  return String(value).toLowerCase() === "true";
}

function propJson_(props, key, fallback) {
  var value = props.getProperty(key);
  if (!value) {
    return fallback;
  }
  try {
    return JSON.parse(value);
  } catch (error) {
    throw new Error(
      key +
        " must be valid JSON. Paste raw JSON without wrapping quotes or backslashes. " +
        "Original error: " +
        error.message
    );
  }
}

function propCsv_(props, key, fallback) {
  var value = props.getProperty(key);
  if (!value) {
    return fallback;
  }
  return value.split(",").map(function(item) {
    return item.trim();
  }).filter(Boolean);
}
