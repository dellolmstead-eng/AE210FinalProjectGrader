import { STRINGS } from "./messages.js";
import { runMissionChecks } from "./rules/mission.js";
import { runAeroChecks } from "./rules/aero.js";
import { runThrustAndTakeoff } from "./rules/thrust.js";
import { runConstraintChecks } from "./rules/constraints.js";
import { runAttachmentChecks } from "./rules/attachments.js";
import { runStealthChecks } from "./rules/stealth.js";
import { runStabilityChecks } from "./rules/stability.js";
import { runFuelVolumeChecks } from "./rules/fuel.js";
import { runRecurringCostChecks } from "./rules/cost.js";
import { runLandingGearChecks } from "./rules/landingGear.js";
import { getCell, asNumber } from "./parseUtils.js";
import { format } from "./format.js";

function cellRef(rowIdx, colIdx) {
  let col = "";
  let n = colIdx;
  while (n >= 0) {
    col = String.fromCharCode((n % 26) + 65) + col;
    n = Math.floor(n / 26) - 1;
  }
  return `${col}${rowIdx + 1}`;
}

const TOLERANCES = {
  tol: 1e-3,
  mach: 1e-2,
  dist: 1e-3,
};

const BONUS_FULL_EPS = 1e-6;
const BONUS_MIN_DISPLAY = 1e-2;

const clamp01 = (value) => Math.max(0, Math.min(1, value));

const roundToTenth = (value) => (Number.isFinite(value) ? Math.round(value * 10) / 10 : 0);

const linearBonus = (value, threshold, objective) => {
  if (!Number.isFinite(value)) return 0;
  if (objective === threshold) return value >= objective ? 1 : 0;
  return clamp01((value - threshold) / (objective - threshold));
};

const inverseLinearBonus = (value, threshold, objective) => {
  if (!Number.isFinite(value)) return 0;
  if (objective === threshold) return value <= objective ? 1 : 0;
  return clamp01((threshold - value) / (threshold - objective));
};

function computeBonuses(workbook) {
  const messages = [];
  let total = 0;

  const pushBonus = (rawPoints, template, ...args) => {
    const points = roundToTenth(rawPoints);
    if (points <= 0) {
      return;
    }
    total += points;
    if (points < 1 - BONUS_FULL_EPS && points >= BONUS_MIN_DISPLAY) {
      messages.push(format(template, points, ...args));
    }
  };

  const main = workbook.sheets.main;

  // Mission radius (linear 0 at 375 nm, 1 at 410 nm)
  const radius = asNumber(getCell(main, "Y37"));
  const radiusBonus = linearBonus(radius, 375, 410);
  pushBonus(radiusBonus, STRINGS.bonus.radius, radius);

  const aim120 = asNumber(getCell(main, "AB3"));
  const aim9 = asNumber(getCell(main, "AB4"));
  const payloadBonus =
    Number.isFinite(aim120) && Number.isFinite(aim9) && aim120 >= 8 - TOLERANCES.tol && aim9 >= 2 - TOLERANCES.tol
      ? 1
      : 0;
  pushBonus(payloadBonus, STRINGS.bonus.payload, aim120 ?? 0, aim9 ?? 0);

  const takeoffDist = asNumber(getCell(main, "X12"));
  const takeoffBonus = inverseLinearBonus(takeoffDist, 3000, 2500);
  pushBonus(takeoffBonus, STRINGS.bonus.takeoff, takeoffDist ?? 0);

  const landingDist = asNumber(getCell(main, "X13"));
  const landingBonus = inverseLinearBonus(landingDist, 5000, 3500);
  pushBonus(landingBonus, STRINGS.bonus.landing, landingDist ?? 0);

  const maxMach = asNumber(getCell(main, "U3"));
  const maxMachBonus = linearBonus(maxMach, 2.0, 2.2);
  pushBonus(maxMachBonus, STRINGS.bonus.maxMach, maxMach ?? 0);

  const supercruiseMach = asNumber(getCell(main, "U4"));
  const supercruiseBonus = linearBonus(supercruiseMach, 1.5, 1.8);
  pushBonus(supercruiseBonus, STRINGS.bonus.supercruise, supercruiseMach ?? 0);

  const psHigh = asNumber(getCell(main, "X8"));
  const psHighBonus = linearBonus(psHigh, 400, 500);
  pushBonus(psHighBonus, STRINGS.bonus.psHigh, psHigh ?? 0);

  const psLow = asNumber(getCell(main, "X9"));
  const psLowBonus = linearBonus(psLow, 400, 500);
  pushBonus(psLowBonus, STRINGS.bonus.psLow, psLow ?? 0);

  const gHigh = asNumber(getCell(main, "V6"));
  const gHighBonus = linearBonus(gHigh, 3.0, 4.0);
  pushBonus(gHighBonus, STRINGS.bonus.gHigh, gHigh ?? 0);

  const gLow = asNumber(getCell(main, "V7"));
  const gLowBonus = linearBonus(gLow, 4.0, 4.5);
  pushBonus(gLowBonus, STRINGS.bonus.gLow, gLow ?? 0);

  const cost = asNumber(getCell(main, "Q31"));
  const numAircraft = asNumber(getCell(main, "N31"));
  let costBonus = 0;
  if (Number.isFinite(cost) && Number.isFinite(numAircraft)) {
    if (Math.abs(numAircraft - 187) < 1e-3) {
      costBonus = inverseLinearBonus(cost, 115, 100);
      pushBonus(costBonus, STRINGS.bonus.cost, numAircraft, cost);
    } else if (Math.abs(numAircraft - 800) < 1e-3) {
      costBonus = inverseLinearBonus(cost, 75, 63);
      pushBonus(costBonus, STRINGS.bonus.cost, numAircraft, cost);
    }
  }

  return { points: roundToTenth(total), messages };
}

export function gradeWorkbook(workbook, rules) {
  // Preflight: abort if Main sheet contains Excel errors / invalid cells
  if (workbook.sheets?.main) {
    const invalidCells = [];
    workbook.sheets.main.forEach((row, rIdx) => {
      if (!row) {
        return;
      }
      row.forEach((value, cIdx) => {
        if (typeof value === "string" && /^#(DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|N\/A)$/i.test(value.trim())) {
          invalidCells.push(cellRef(rIdx, cIdx));
        } else if (typeof value === "number" && !Number.isFinite(value)) {
          invalidCells.push(cellRef(rIdx, cIdx));
        }
      });
    });
    if (invalidCells.length > 0) {
      const msg = `Invalid for analysis: Excel errors in Main sheet at ${invalidCells.join(", ")}. Correct the errors and resubmit.`;
      return {
        score: 0,
        maxScore: 40,
        scoreLine: msg,
        bonusLine: "",
        feedbackLog: msg,
      };
    }
  }

  const feedback = [];
  let baseScore = 40;

  if (workbook.fileName) {
    feedback.push(workbook.fileName);
  }

  const aeroResult = runAeroChecks(workbook, rules);
  baseScore += aeroResult.delta;
  feedback.push(...aeroResult.feedback);

  const missionResult = runMissionChecks(workbook, rules);
  baseScore += missionResult.delta;
  feedback.push(...missionResult.feedback);

  const thrustResult = runThrustAndTakeoff(workbook, rules);
  baseScore += thrustResult.delta;
  feedback.push(...thrustResult.feedback);

  const attachmentResult = runAttachmentChecks(workbook, rules);
  baseScore += attachmentResult.delta;
  feedback.push(...attachmentResult.feedback);

  const stealthResult = runStealthChecks(workbook);
  baseScore += stealthResult.delta;
  feedback.push(...stealthResult.feedback);

  const constraintResult = runConstraintChecks(workbook, rules);
  baseScore += constraintResult.delta;
  baseScore += constraintResult.payloadDelta;
  baseScore += constraintResult.curveDelta;
  feedback.push(...constraintResult.feedback);

  const stabilityResult = runStabilityChecks(workbook, rules);
  baseScore += stabilityResult.delta;
  feedback.push(...stabilityResult.feedback);

  const fuelResult = runFuelVolumeChecks(workbook, rules);
  baseScore += fuelResult.delta;
  feedback.push(...fuelResult.feedback);

  const costResult = runRecurringCostChecks(workbook, rules);
  baseScore += costResult.delta;
  feedback.push(...costResult.feedback);

  const gearResult = runLandingGearChecks(workbook, rules);
  baseScore += gearResult.delta;
  feedback.push(...gearResult.feedback);

  const clampedBase = Math.max(0, baseScore);

  const bonusResult = computeBonuses(workbook);
  feedback.push(...bonusResult.messages);
  const finalScore = roundToTenth(clampedBase + bonusResult.points);

  const scoreLine = format(STRINGS.summary.base, clampedBase);
  const bonusLine = format(STRINGS.summary.bonus, bonusResult.points, finalScore);

  feedback.push(scoreLine);
  feedback.push(bonusLine);

  return {
    score: finalScore,
    maxScore: 40,
    scoreLine,
    bonusLine,
    feedbackLog: feedback.join("\n"),
  };
}
