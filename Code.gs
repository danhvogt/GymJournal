function doGet(e) {
  if (e && e.parameter && e.parameter.page == "progress") {
    return HtmlService.createHtmlOutputFromFile("progress");
  }
  return HtmlService.createHtmlOutputFromFile("index");
}

// Get exercises by type
function getExercises(type) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exercises");
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(header => String(header).trim().toLowerCase());
  const rows = data.slice(1);

  const typeIndex = headers.indexOf("type");
  const exerciseIndex = headers.indexOf("exercise");
  const sortOrderIndex = headers.indexOf("sortorder");
  const gifIndex = headers.indexOf("gif");

  return rows
    .filter(row => row[typeIndex] === type)
    .sort((a, b) => a[sortOrderIndex] - b[sortOrderIndex])
    .map(row => ({
      name: row[exerciseIndex],
      gif: gifIndex >= 0 ? normalizeGifUrl(row[gifIndex]) : ""
    }));
}

function normalizeGifUrl(value) {
  const url = String(value || "").trim();

  if (!url) return "";

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

// Get last successful workout
function getLastWorkout(exercise) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Log");
  const data = sheet.getDataRange().getValues().slice(1).reverse();

  let lastSuccess = null;
  let lastAttempt = null;

  for (let row of data) {

    if (row[2] === exercise) {

      // first match = most recent attempt
      if (!lastAttempt) {
        lastAttempt = {
          sets: row[3],
          reps: row[4],
          load: row[5],
          notes: row[7]
        };
      }

      // first successful match
      if (!lastSuccess && row[6] === true) {
        lastSuccess = {
          sets: row[3],
          reps: row[4],
          load: row[5]
        };
      }

      // stop early if we have both
      if (lastAttempt && lastSuccess) break;
    }
  }

  return {
    success: lastSuccess,
    attempt: lastAttempt
  };
}

// Save new workout
function saveWorkout(entry) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Log");

  sheet.appendRow([
    new Date(),
    entry.type,
    entry.exercise,
    entry.sets,
    entry.reps,
    entry.load,
    entry.success,
    entry.notes   // <-- ADD THIS
  ]);

  return true;
}

function getProgressDataByType(type, successfulOnly) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Log");
  const data = sheet.getDataRange().getValues().slice(1);
  const onlySuccessful = successfulOnly === true;

  // filter by type
  const filtered = data
    .filter(row => row[1] === type)
    .filter(row => !onlySuccessful || row[6] === true)
    .sort((a, b) => new Date(a[0]) - new Date(b[0]));

  // group by exercise
  const grouped = {};

  filtered.forEach(row => {
    const exercise = row[2];
    const sets = Number(row[3]);
    const reps = Number(row[4]);
    const load = Number(row[5]);
    const success = row[6] === true;
    const volume = sets * reps * load;
    const estimated1rm = load > 0 && reps > 0
      ? Math.round((load * (1 + reps / 30)) * 10) / 10
      : 0;

    if (!grouped[exercise]) {
      grouped[exercise] = [];
    }

    grouped[exercise].push({
      date: row[0],
      sets: sets,
      reps: reps,
      load: load,
      volume: volume,
      estimated1rm: estimated1rm,
      success: success
    });
  });

  // convert into chart-ready structure
  const result = [];

  Object.keys(grouped).forEach(exercise => {
    const workouts = grouped[exercise];

    result.push({
      exercise: exercise,
      labels: workouts.map((_, i) => i + 1),
      workoutCount: workouts.length,
      metrics: {
        volume: workouts.map(item => item.volume),
        load: workouts.map(item => item.load),
        reps: workouts.map(item => item.reps),
        estimated1rm: workouts.map(item => item.estimated1rm)
      }
    });
  });

  return result;
}
