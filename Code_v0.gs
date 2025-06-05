// --- Configuration ---
const FOLDER_ID = '1f7esiYvQj7MElui8SFYnJ8Vc7GmBlQ5G'; // <-- Replace with your Drive folder ID

// !!! SECURITY WARNING !!!
// Hardcoding API keys directly in code is insecure, especially if the script
// is shared or part of a larger project. Consider using Script Properties
// with User scope or OAuth 2.0 for better security.
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'); //

const ALERT_EMAIL = 'kshitij@prisonopticians.org'; // <-- Replace with the email address for alerts
const SCRIPT_PROP_LAST_CHECK = 'lastCheckTimestamp';

// New: Retry configuration for API calls
const MAX_RETRIES = 15; // Max number of retries for a failed API call (including initial attempt)
const INITIAL_DELAY_MS = 1000; // Initial delay before first retry (1 second)

// Tolerance settings (Based on ANSI Z80.1-2020 Standards - Verify/Adjust as needed)
const TOLERANCE = {
  SPHERE: 0.13, // Diopters (for |S| <= 6.50 D)
  CYLINDER: 0.13, // Diopters
  AXIS: { // Cylinder Axis Tolerance in Degrees
    cyl_le_025: 14,  // For |C| <= 0.25 D
    cyl_le_050: 7,   // For 0.25 D < |C| <= 0.50 D
    cyl_le_075: 5,   // For 0.50 D < |C| <= 0.75 D
    cyl_le_150: 3,   // For 0.75 D < |C| <= 1.50 D
    cyl_gt_150: 2    // For |C| > 1.50 D
  },
  ADD: 0.12,      // Add Power Tolerance in Diopters
  PRISM_H: 0.33,  // Horizontal Prism Tolerance per lens (Δ)
  PRISM_V: 0.33   // Vertical Prism Tolerance per lens (Δ)
  // Note: ANSI also specifies imbalance tolerances (R vs L), not checked here directly
};

// Gemini API endpoint
const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;
//const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key=${GEMINI_API_KEY}`;

// --- Main Function to Check for New Files ---
function checkNewImagesAndQC() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const properties = PropertiesService.getScriptProperties();
  // Get the last check time; default to 0 if not set
  const lastCheckTimestamp = parseInt(properties.getProperty(SCRIPT_PROP_LAST_CHECK) || '0');
  const nowTimestamp = new Date().getTime();

  // Ensure we don't check the exact same second again if run frequently
  const checkFromTimestamp = lastCheckTimestamp + 1;

  Logger.log(`Checking for files modified after: ${new Date(checkFromTimestamp).toISOString()}`);

  // IMPORTANT: The search query is currently 'modifiedDate < ...'
  // For production, you almost certainly want 'modifiedDate > "' + new Date(checkFromTimestamp).toISOString() + '"'
  // I've kept your temporary change for now, but ensure this is correct for deployment.
  const files = folder.searchFiles(
    'modifiedDate < "' + new Date(checkFromTimestamp).toISOString() + '" and (mimeType = "image/jpeg" or mimeType = "image/png")'
  ); // ## change < to > later, at deployment

  let processedFiles = 0;
  while (files.hasNext()) {
    const file = files.next();
    // Double check modification time just in case searchFiles is slightly delayed
    // Re-enable this for production to prevent re-processing older files after initial testing
    // if (file.getLastUpdated().getTime() < checkFromTimestamp) {
    //    Logger.log(`Skipping file ${file.getName()} as its modification time (${file.getLastUpdated()}) is before the last check time.`);
    //    continue;
    // }

    Logger.log(`Processing file: ${file.getName()} (ID: ${file.getId()}, Modified: ${file.getLastUpdated()})`);
    try {
      processImageWithGemini(file);
      processedFiles++;
    } catch (e) {
      Logger.log(`Error processing file ${file.getName()}: ${e}\nStack: ${e.stack}`);
      sendAlert(`QC Script Error Processing ${file.getName()}`, `File: ${file.getName()}\nURL: ${file.getUrl()}\nError: ${e}\nStack: ${e.stack}`);
    }
  }

  // Update the last check timestamp
  properties.setProperty(SCRIPT_PROP_LAST_CHECK, nowTimestamp.toString());
  Logger.log(`Checked folder. Processed ${processedFiles} new files. Last check updated to: ${new Date(nowTimestamp)}`);
}

// --- Function to Process a Single Image ---
function processImageWithGemini(file) {
  const fileBlob = file.getBlob();
  const base64Image = Utilities.base64Encode(fileBlob.getBytes());

  // Construct the prompt for Gemini - UPDATED FOR COMPLEX PRESCRIPTIONS
  const prompt = `
    Analyze the provided image showing an eyeglass prescription slip and a lensmeter measurement screen.
    Extract the following values for BOTH the Right (R) and Left (L) eyes from BOTH the prescription ("prescription") and the measurement screen ("measurement"):
    1. Sphere (S): Numerical value (e.g., +1.50, -2.75, 0).
    2. Cylinder (C): Numerical value (e.g., -0.50, 0). Use negative cylinder convention if possible.
    3. Axis (A): Integer value (0-180). Use 0 if Cylinder is 0 or not present.
    4. Add Power (ADD): Numerical value (e.g., +2.00, +2.50). Use 0 or null if not a bifocal/progressive or not specified.
    5. Horizontal Prism Magnitude (PrismH): Numerical value (prism diopters, e.g., 1.00, 0.50). Use 0 or null if no horizontal prism.
    6. Vertical Prism Magnitude (PrismV): Numerical value (prism diopters, e.g., 2.00, 0.33). Use 0 or null if no vertical prism.
    7. Horizontal Base Direction (PrismBaseH): String 'IN' or 'OUT'. Use null if PrismH is 0.
    8. Vertical Base Direction (PrismBaseV): String 'UP' or 'DOWN'. Use null if PrismV is 0.

    IMPORTANT: Structure the output STRICTLY as a single JSON object with top-level keys "prescription" and "measurement".
    Each of these keys should contain an object with keys "R" and "L".
    Each "R" and "L" key should contain an object with ALL the keys: "S", "C", "A", "ADD", "PrismH", "PrismV", "PrismBaseH", "PrismBaseV".
    Ensure numerical fields are numbers, axis is an integer, and base directions are strings ('IN', 'OUT', 'UP', 'DOWN') or null.
    Top row in prescription is Right eye and bottom row is left eye.
    Right side of the blue screen (measurement) is right eye measurement and left side is the left side measurement.
    Do not add any comments in the json file.
    Do not add + sign in response for positive prescription numbers.

    Example JSON format:
    {
      "prescription": {
        "R": { "S": -2.75, "C": -0.50, "A": 90, "ADD": 2.50, "PrismH": 1.00, "PrismV": 0, "PrismBaseH": "OUT", "PrismBaseV": null },
        "L": { "S": -3.25, "C": -0.50, "A": 70, "ADD": 2.50, "PrismH": 1.00, "PrismV": 0, "PrismBaseH": "OUT", "PrismBaseV": null }
      },
      "measurement": {
        "R": { "S": -2.75, "C": -0.50, "A": 95, "ADD": 2.50, "PrismH": 1.15, "PrismV": 0.34, "PrismBaseH": "OUT", "PrismBaseV": "UP" },
        "L": { "S": -3.25, "C": -0.50, "A": 76, "ADD": 2.50, "PrismH": 1.24, "PrismV": 0.63, "PrismBaseH": "IN", "PrismBaseV": "UP" }
      }
    }
    `;

  const payload = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: fileBlob.getContentType(), data: base64Image } }] }],
    // Lower temperature might give more consistent JSON structure
    generationConfig: { "temperature": 0.1, "maxOutputTokens": 2048 }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let attempt = 0;
  let delay = INITIAL_DELAY_MS;
  let response;
  let responseCode;
  let responseBody = ''; // Initialize responseBody to empty string

  // Loop for retries
  while (attempt < MAX_RETRIES) {
    Logger.log(`Attempt ${attempt + 1}/${MAX_RETRIES} for file ${file.getName()}`);
    try {
      response = UrlFetchApp.fetch(GEMINI_API_URL, options);
      responseCode = response.getResponseCode();
      responseBody = response.getContentText(); // Update responseBody after each fetch

      // Check for success (200 OK)
      if (responseCode === 200) {
        Logger.log(`Gemini API call successful for ${file.getName()} on attempt ${attempt + 1}.`);
        break; // Exit retry loop on success
      }
      // Check for retryable errors (429 Too Many Requests or any 5xx Server Error)
      else if (responseCode === 429 || (responseCode >= 500 && responseCode < 600)) {
        Logger.log(`Gemini API returned retryable error ${responseCode} for ${file.getName()}. Retrying in ${delay / 1000} seconds...`);
        Utilities.sleep(delay);
        delay *= 2; // Exponential backoff
        attempt++;
        if (attempt === MAX_RETRIES) {
            // Log that max retries are reached before breaking/throwing
            Logger.log(`Max retries reached for file ${file.getName()} with error ${responseCode}.`);
        }
      }
      // For non-retryable errors (e.g., 400 Bad Request, 401 Unauthorized, 403 Forbidden)
      else {
        const errorMessage = `Non-retryable Gemini API error for file ${file.getName()}. Code: ${responseCode}. Response: ${responseBody}`;
        Logger.log(errorMessage);
        sendAlert(`QC Error: Non-Retryable Gemini API Error (${responseCode})`, `File: ${file.getName()}\nURL: ${file.getUrl()}\nError: ${errorMessage.substring(0,1000)}`);
        throw new Error(errorMessage); // Fail immediately for non-retryable errors
      }
    } catch (e) {
      // Catch network errors (e.g., DNS issues, connection problems) or other unexpected errors
      Logger.log(`Network or unexpected error during Gemini API call for ${file.getName()} on attempt ${attempt + 1}: ${e}\nRetrying in ${delay / 1000} seconds...`);
      Utilities.sleep(delay);
      delay *= 2; // Exponential backoff
      attempt++;
      if (attempt === MAX_RETRIES) {
          Logger.log(`Max retries reached for file ${file.getName()} due to network/unexpected error.`);
      }
    }
  }

  // After the retry loop, check if the request was successful
  if (responseCode !== 200) {
    const finalErrorMessage = `Failed to get successful response from Gemini API for file ${file.getName()} after ${MAX_RETRIES} attempts. Last code: ${responseCode || 'N/A'}, Last Response: ${responseBody.substring(0, 500)}`;
    Logger.log(finalErrorMessage);
    sendAlert(`QC Error: Gemini API Retries Exhausted`, `File: ${file.getName()}\nURL: ${file.getUrl()}\nError: ${finalErrorMessage}`);
    throw new Error(finalErrorMessage); // Re-throw to indicate failure for this file
  }

  // --- Process the successful 200 response ---
  try {
    const jsonResponse = JSON.parse(responseBody);
    if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
        jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
        jsonResponse.candidates[0].content.parts.length > 0) { // Fixed: jsonResponse.candidates[0]

        let extractedText = jsonResponse.candidates[0].content.parts[0].text;
        Logger.log(`Gemini Raw Response Text for ${file.getName()}: ${extractedText}`);

        // Clean potential markdown formatting
        extractedText = extractedText.replace(/```json/g, '').replace(/```/g, '').trim();

        // Validate and parse the JSON
        const data = JSON.parse(extractedText);

        // Basic validation of structure
        if (!data || !data.prescription || !data.measurement ||
            !data.prescription.R || !data.prescription.L ||
            !data.measurement.R || !data.measurement.L) {
           throw new Error("Parsed JSON missing required structure (prescription/measurement or R/L keys).");
        }

        Logger.log(`Parsed Data for ${file.getName()}: ${JSON.stringify(data)}`);
        performQC(data.prescription, data.measurement, file.getName(), file.getUrl());

    } else if (jsonResponse.promptFeedback && jsonResponse.promptFeedback.blockReason) {
        // Handle cases where the prompt was blocked
        const reason = jsonResponse.promptFeedback.blockReason;
        const safetyRatings = JSON.stringify(jsonResponse.promptFeedback.safetyRatings || {});
        Logger.log(`Gemini request blocked for file ${file.getName()}. Reason: ${reason}. Ratings: ${safetyRatings}`);
        sendAlert(`QC Error: Gemini Request Blocked (${reason})`, `File: ${file.getName()}\nURL: ${file.getUrl()}\nReason: ${reason}\nSafety Ratings: ${safetyRatings}`);
    } else {
        // Handle other unexpected valid JSON responses
        Logger.log(`Error: Unexpected Gemini response structure for file ${file.getName()}. Response: ${responseBody}`);
        sendAlert(`QC Error: Unexpected Gemini response structure`, `File: ${file.getName()}\nURL: ${file.getUrl()}\nResponse: ${responseBody.substring(0, 1000)}`);
    }
  } catch (e) {
    Logger.log(`Error parsing Gemini response for file ${file.getName()}: ${e}\nResponse Body: ${responseBody}`);
    sendAlert(`QC Error: Cannot Parse Gemini Response`, `File: ${file.getName()}\nURL: ${file.getUrl()}\nError: ${e}\nRaw Response Text: ${responseBody.substring(0, 1000)}`);
  }
}

// --- Function to Compare Data and Send Alert ---
function performQC(prescription, measurement, fileName, fileUrl) {
  let mismatches = [];

  // Compare Right Eye
  compareEye('R', prescription.R, measurement.R, mismatches);
  // Compare Left Eye
  compareEye('L', prescription.L, measurement.L, mismatches);

  if (mismatches.length > 0) {
    let alertBody = `QC Alert: Mismatch detected for file: ${fileName}\n`;
    alertBody += `File URL: ${fileUrl}\n\n`;
    alertBody += `Mismatches found:\n`;
    mismatches.forEach(m => {
      alertBody += `- ${m}\n`;
    });

    // --- Append detailed comparison data to email ---
    try {
        alertBody += `\n--- Details ---\n`;
        alertBody += `Prescription R: ${formatEyeData(prescription.R)}\n`;
        alertBody += `Measurement R:  ${formatEyeData(measurement.R)}\n`;
        alertBody += `Prescription L: ${formatEyeData(prescription.L)}\n`;
        alertBody += `Measurement L:  ${formatEyeData(measurement.L)}\n`;
    } catch(e) {
        Logger.log(`Error formatting details for alert: ${e}`);
        // Fallback to raw JSON if formatting fails
        alertBody += `\nPrescription:\n R: ${JSON.stringify(prescription.R)}\n L: ${JSON.stringify(prescription.L)}\n`;
        alertBody += `\nMeasurement:\n R: ${JSON.stringify(measurement.R)}\n L: ${JSON.stringify(measurement.L)}\n`;
    }
    // --- End appended details ---

    Logger.log(`Sending QC Alert for ${fileName}`);
    sendAlert(`QC Alert: Mismatch for ${fileName}`, alertBody);
  } else {
    Logger.log(`QC Passed for file: ${fileName}`);
    // Optional: Log success or take other action (e.g., move file to a "Passed" folder)
  }
}

// --- Helper function to format eye data for alerts ---
function formatEyeData(eyeData) {
    if (!eyeData) return "N/A";
    let parts = [];
    if (eyeData.S !== undefined && eyeData.S !== null) parts.push(`S: ${eyeData.S.toFixed(2)}`);
    if (eyeData.C !== undefined && eyeData.C !== null) parts.push(`C: ${eyeData.C.toFixed(2)}`);
    if (eyeData.A !== undefined && eyeData.A !== null) parts.push(`A: ${eyeData.A}`);
    if (eyeData.ADD !== undefined && eyeData.ADD !== null && eyeData.ADD != 0) parts.push(`Add: ${eyeData.ADD.toFixed(2)}`);
    if (eyeData.PrismH !== undefined && eyeData.PrismH !== null && eyeData.PrismH != 0) {
        parts.push(`Prism H: ${eyeData.PrismH.toFixed(2)} ${eyeData.PrismBaseH || ''}`);
    }
    if (eyeData.PrismV !== undefined && eyeData.PrismV !== null && eyeData.PrismV != 0) {
        parts.push(`Prism V: ${eyeData.PrismV.toFixed(2)} ${eyeData.PrismBaseV || ''}`);
    }
    return parts.join(', ') || "Data missing";
}

// --- Helper function to compare one eye ---
function compareEye(eye, pres, meas, mismatches) {
    // --- Safely parse all values, defaulting to 0 or null ---
    const presS = parseFloat(pres?.S ?? 0);
    const presC = parseFloat(pres?.C ?? 0);
    const presA = parseInt(pres?.A ?? 0);
    const presAdd = parseFloat(pres?.ADD ?? 0);
    const presPrismH = Math.abs(parseFloat(pres?.PrismH ?? 0)); // Use magnitude for comparison
    const presPrismV = Math.abs(parseFloat(pres?.PrismV ?? 0)); // Use magnitude for comparison
    const presBaseH = pres?.PrismBaseH?.toUpperCase() || null;
    const presBaseV = pres?.PrismBaseV?.toUpperCase() || null;

    const measS = parseFloat(meas?.S ?? 0);
    const measC = parseFloat(meas?.C ?? 0);
    let measA = parseInt(meas?.A ?? 0);
    const measAdd = parseFloat(meas?.ADD ?? 0);
    const measPrismH = Math.abs(parseFloat(meas?.PrismH ?? 0)); // Use magnitude for comparison
    const measPrismV = Math.abs(parseFloat(meas?.PrismV ?? 0)); // Use magnitude for comparison
    const measBaseH = meas?.PrismBaseH?.toUpperCase() || null;
    const measBaseV = meas?.PrismBaseV?.toUpperCase() || null;

    // Helper to check if a value is effectively zero (within tolerance)
    // Using tolerance value itself as the threshold for zero check related to prism presence/absence
    const isEffectivelyZero = (val, tolerance = 0.01) => Math.abs(val) < tolerance;
    const isPrismEffectivelyZeroH = (val) => Math.abs(val) <= TOLERANCE.PRISM_H;
    const isPrismEffectivelyZeroV = (val) => Math.abs(val) <= TOLERANCE.PRISM_V;

    // Helper to check base direction mismatch (e.g., 'UP' vs 'DOWN')
    const isOppositeBase = (base1, base2) => {
        return (base1 === 'UP' && base2 === 'DOWN') || (base1 === 'DOWN' && base2 === 'UP') ||
               (base1 === 'IN' && base2 === 'OUT') || (base1 === 'OUT' && base2 === 'IN');
    };

    // Normalize axis if cylinder is zero
    if (isEffectivelyZero(measC) && measA === 180) measA = 0;
    const finalPresA = isEffectivelyZero(presC) ? 0 : presA;
    const finalMeasA = isEffectivelyZero(measC) ? 0 : measA;

    // --- Comparisons ---

    // Sphere
    if (Math.abs(presS - measS) > TOLERANCE.SPHERE) {
        mismatches.push(`${eye} Eye Sphere Mismatch: Prescription=${presS.toFixed(2)}, Measured=${measS.toFixed(2)} (Tolerance: +/-${TOLERANCE.SPHERE})`);
    }

    // Cylinder
    if (Math.abs(presC - measC) > TOLERANCE.CYLINDER) {
        mismatches.push(`${eye} Eye Cylinder Mismatch: Prescription=${presC.toFixed(2)}, Measured=${measC.toFixed(2)} (Tolerance: +/-${TOLERANCE.CYLINDER})`);
    }

    // Axis (only if prescription cylinder is significant)
    const presCIsSignificant = !isEffectivelyZero(presC, 0.01); // Use small epsilon for significance check
    if (presCIsSignificant) {
        const absPresC = Math.abs(presC);
        let axisTolerance;
        if (absPresC <= 0.25) axisTolerance = TOLERANCE.AXIS.cyl_le_025;
        else if (absPresC <= 0.50) axisTolerance = TOLERANCE.AXIS.cyl_le_050;
        else if (absPresC <= 0.75) axisTolerance = TOLERANCE.AXIS.cyl_le_075;
        else if (absPresC <= 1.50) axisTolerance = TOLERANCE.AXIS.cyl_le_150;
        else axisTolerance = TOLERANCE.AXIS.cyl_gt_150;

        const diff = Math.abs(finalPresA - finalMeasA);
        const axisDiff = Math.min(diff, 180 - diff); // Shortest angle difference

        if (axisDiff > axisTolerance) {
            mismatches.push(`${eye} Eye Axis Mismatch: Prescription=${finalPresA}, Measured=${finalMeasA} (Tolerance: +/-${axisTolerance}, Diff: ${axisDiff.toFixed(1)})`);
        }
    } else if (!isEffectivelyZero(measC, TOLERANCE.CYLINDER)) {
        // Case: Prescription has effectively zero cylinder, but measurement has significant cylinder
        mismatches.push(`${eye} Eye Cylinder Mismatch: Prescription=0.00, Measured=${measC.toFixed(2)} (Tolerance: +/-${TOLERANCE.CYLINDER})`);
    }

    // Add Power (only if prescription Add is significant)
     const presAddIsSignificant = !isEffectivelyZero(presAdd, 0.01);
    if (presAddIsSignificant) {
        if (Math.abs(presAdd - measAdd) > TOLERANCE.ADD) {
            mismatches.push(`${eye} Eye Add Power Mismatch: Prescription=${presAdd.toFixed(2)}, Measured=${measAdd.toFixed(2)} (Tolerance: +/-${TOLERANCE.ADD})`);
        }
    } else if (!isEffectivelyZero(measAdd, TOLERANCE.ADD)) {
         // Case: Prescription has no add, but measurement has significant add
        mismatches.push(`${eye} Eye Add Power Mismatch: Prescription=0.00, Measured=${measAdd.toFixed(2)} (Tolerance: +/-${TOLERANCE.ADD})`);
    }

    // --- MODIFIED PRISM LOGIC ---
    // Check prism only if it's significantly prescribed

    const isPrescribedPrismH = presPrismH > 0.01; // Check if prescribed H prism is intended (not just noise)
    const isPrescribedPrismV = presPrismV > 0.01; // Check if prescribed V prism is intended

    // Horizontal Prism Check (Only perform checks if prescribed)
    if (isPrescribedPrismH) {
        // Check Magnitude Difference
        if (Math.abs(presPrismH - measPrismH) > TOLERANCE.PRISM_H) {
            mismatches.push(`${eye} Eye Horizontal Prism Magnitude Mismatch: Prescribed=${presPrismH.toFixed(2)}, Measured=${measPrismH.toFixed(2)} (Tolerance: +/-${TOLERANCE.PRISM_H})`);
        }
        // Check Base Direction (only if measured prism is also significant enough to have a reliable base)
        else if (measPrismH > TOLERANCE.PRISM_H && presBaseH && measBaseH && presBaseH !== measBaseH) {
             mismatches.push(`${eye} Eye Horizontal Prism Base Mismatch: Prescribed=${presBaseH}, Measured=${measBaseH}`);
        }
        // Check if prescribed prism is missing in measurement (measured is effectively zero)
        else if (isPrismEffectivelyZeroH(measPrismH)) {
             mismatches.push(`${eye} Eye Missing Horizontal Prism: Prescribed=${presPrismH.toFixed(2)} ${presBaseH || ''}, Measured near zero`);
        }
    }
    // *** Removed check for unexpected measured H prism when none is prescribed ***

    // Vertical Prism Check (Only perform checks if prescribed)
    if (isPrescribedPrismV) {
        // Check Magnitude Difference
        if (Math.abs(presPrismV - measPrismV) > TOLERANCE.PRISM_V) {
             mismatches.push(`${eye} Eye Vertical Prism Magnitude Mismatch: Prescribed=${presPrismV.toFixed(2)}, Measured=${measPrismV.toFixed(2)} (Tolerance: +/-${TOLERANCE.PRISM_V})`);
        }
        // Check Base Direction (especially opposite, only if measured is significant enough)
        else if (measPrismV > TOLERANCE.PRISM_V && presBaseV && measBaseV && presBaseV !== measBaseV) {
            if (isOppositeBase(presBaseV, measBaseV)) {
                mismatches.push(`${eye} Eye Vertical Prism Base Mismatch (Opposite): Prescribed=${presBaseV}, Measured=${measBaseV}`);
            } else {
                 // Optional: Flag non-opposite mismatches too if needed
                 mismatches.push(`${eye} Eye Vertical Prism Base Mismatch: Prescribed=${presBaseV}, Measured=${measBaseV}`);
            }
        }
        // Check if prescribed prism is missing in measurement (measured is effectively zero)
        else if (isPrismEffectivelyZeroV(measPrismV)) {
             mismatches.push(`${eye} Eye Missing Vertical Prism: Prescribed=${presPrismV.toFixed(2)} ${presBaseV || ''}, Measured near zero`);
        }
    }
    // *** Removed check for unexpected measured V prism when none is prescribed ***
} // --- End of compareEye function ---


// --- Helper function to send email alerts ---
function sendAlert(subject, body) {
  try {
      MailApp.sendEmail(ALERT_EMAIL, subject, body);
  } catch (e) {
      Logger.log(`Failed to send email alert. Subject: ${subject}. Error: ${e}`);
      // Consider alternative notification if email fails consistently
  }
}
