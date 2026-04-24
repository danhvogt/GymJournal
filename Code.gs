function doGet(e) {
  if (e && e.parameter && e.parameter.page == "progress") {
    return HtmlService.createHtmlOutputFromFile("progress");
  }

  return HtmlService.createHtmlOutputFromFile("index");
}

function getExercises(type) {
  const rows = getTableRows_("Exercises");
  const selectedType = String(type || "").trim();

  return rows
    .filter(row => row.type === selectedType)
    .sort((a, b) => a.sortOrder - b.sortOrder)
    .map(row => ({
      name: row.exercise,
      gif: normalizeGifUrl(row.gif)
    }));
}

function getLastWorkout(exercise) {
  const selectedExercise = String(exercise || "").trim();
  const data = getLogRows_().reverse();

  let lastSuccess = null;
  let lastAttempt = null;

  for (const row of data) {
    if (row.exercise !== selectedExercise) {
      continue;
    }

    if (!lastAttempt) {
      lastAttempt = {
        sets: row.sets,
        reps: row.reps,
        load: row.load,
        notes: row.notes
      };
    }

    if (!lastSuccess && row.success === true) {
      lastSuccess = {
        sets: row.sets,
        reps: row.reps,
        load: row.load
      };
    }

    if (lastAttempt && lastSuccess) {
      break;
    }
  }

  return {
    success: lastSuccess,
    attempt: lastAttempt
  };
}

function saveWorkout(entry) {
  const sheet = getRequiredSheet_("Log");

  sheet.appendRow([
    new Date(),
    String(entry.type || "").trim(),
    String(entry.exercise || "").trim(),
    Number(entry.sets || 0),
    Number(entry.reps || 0),
    Number(entry.load || 0),
    entry.success === true,
    String(entry.notes || "")
  ]);

  return true;
}

function getProgressDataByType(type, successfulOnly) {
  const selectedType = String(type || "").trim();
  const onlySuccessful = successfulOnly === true;
  const grouped = {};

  getLogRows_()
    .filter(row => row.type === selectedType)
    .filter(row => !onlySuccessful || row.success === true)
    .sort((a, b) => new Date(a.date) - new Date(b.date))
    .forEach(row => {
      const volume = row.sets * row.reps * row.load;
      const estimated1rm = row.load > 0 && row.reps > 0
        ? Math.round((row.load * (1 + row.reps / 30)) * 10) / 10
        : 0;

      if (!grouped[row.exercise]) {
        grouped[row.exercise] = [];
      }

      grouped[row.exercise].push({
        date: row.date,
        sets: row.sets,
        reps: row.reps,
        load: row.load,
        volume: volume,
        estimated1rm: estimated1rm,
        success: row.success
      });
    });

  return Object.keys(grouped).map(exercise => {
    const workouts = grouped[exercise];

    return {
      exercise: exercise,
      labels: workouts.map((_, index) => index + 1),
      workoutCount: workouts.length,
      metrics: {
        volume: workouts.map(item => item.volume),
        load: workouts.map(item => item.load),
        reps: workouts.map(item => item.reps),
        estimated1rm: workouts.map(item => item.estimated1rm)
      }
    };
  });
}

function getTableRows_(sheetName) {
  const sheet = getRequiredSheet_(sheetName);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0].map(header => normalizeHeader_(header));
  const typeIndex = headers.indexOf("type");
  const exerciseIndex = headers.indexOf("exercise");
  const sortOrderIndex = headers.indexOf("sortorder");
  const gifIndex = headers.indexOf("gif");

  if (typeIndex < 0 || exerciseIndex < 0) {
    throw new Error("Exercises sheet must include Type and Exercise headers.");
  }

  return data.slice(1)
    .map((row, index) => ({
      type: String(row[typeIndex] || "").trim(),
      exercise: String(row[exerciseIndex] || "").trim(),
      sortOrder: sortOrderIndex >= 0 && row[sortOrderIndex] !== ""
        ? Number(row[sortOrderIndex])
        : index + 1,
      gif: gifIndex >= 0 ? String(row[gifIndex] || "").trim() : ""
    }))
    .filter(row => row.exercise);
}

function getLogRows_() {
  const sheet = getRequiredSheet_("Log");
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0].map(header => normalizeHeader_(header));
  const dateIndex = findHeader_(headers, ["date", "timestamp"]);
  const typeIndex = findHeader_(headers, ["type"]);
  const exerciseIndex = findHeader_(headers, ["exercise"]);
  const setsIndex = findHeader_(headers, ["sets"]);
  const repsIndex = findHeader_(headers, ["reps"]);
  const loadIndex = findHeader_(headers, ["load", "weight"]);
  const successIndex = findHeader_(headers, ["success", "successful"]);
  const notesIndex = findOptionalHeader_(headers, ["notes", "note"]);

  return data.slice(1)
    .filter(row => String(row[exerciseIndex] || "").trim())
    .map(row => ({
      date: row[dateIndex],
      type: String(row[typeIndex] || "").trim(),
      exercise: String(row[exerciseIndex] || "").trim(),
      sets: Number(row[setsIndex] || 0),
      reps: Number(row[repsIndex] || 0),
      load: Number(row[loadIndex] || 0),
      success: normalizeBoolean_(row[successIndex]),
      notes: notesIndex >= 0 ? String(row[notesIndex] || "") : ""
    }));
}

function getRequiredSheet_(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("Missing sheet: " + sheetName);
  }

  return sheet;
}

function normalizeHeader_(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, "");
}

function findHeader_(headers, names) {
  for (const name of names) {
    const index = headers.indexOf(name);

    if (index >= 0) {
      return index;
    }
  }

  throw new Error("Missing required column: " + names[0]);
}

function findOptionalHeader_(headers, names) {
  for (const name of names) {
    const index = headers.indexOf(name);

    if (index >= 0) {
      return index;
    }
  }

  return -1;
}

function normalizeBoolean_(value) {
  if (value === true || value === false) {
    return value;
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  return String(value || "").trim().toLowerCase() === "true";
}

function normalizeGifUrl(value) {
  const url = String(value || "").trim();

  if (!url) {
    return "";
  }

  const driveFileId = extractDriveFileId(url);

  if (driveFileId) {
    return "https://drive.google.com/uc?export=view&id=" + driveFileId;
  }

  return url;
}

function extractDriveFileId(url) {
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/uc\?(?:[^#]*[&?])?id=([a-zA-Z0-9_-]+)/,
    /\/open\?(?:[^#]*[&?])?id=([a-zA-Z0-9_-]+)/
  ];

  for (const pattern of patterns) {
    const match = String(url).match(pattern);

    if (match && match[1]) {
      return match[1];
    }
  }

  return "";
}
