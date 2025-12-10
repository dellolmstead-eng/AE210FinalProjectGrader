import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const PERCENT_TOL = 0.5; // percent span (matches MATLAB GE3 logic)
const ANGLE_TOL = 0.1;
const SPEED_TOL = 0.5;

export function runLandingGearChecks(workbook) {
  const feedback = [];
  let failures = 0;

  const gear = workbook.sheets.gear;

  const noseRule = asNumber(getCell(gear, "J19"));
  if (!Number.isFinite(noseRule) || noseRule < 10 - PERCENT_TOL || noseRule > 20 + PERCENT_TOL) {
    feedback.push(format(STRINGS.gear.nose, noseRule));
    failures += 1;
  }

  const tipbackUpper = asNumber(getCell(gear, "L20"));
  const tipbackLower = asNumber(getCell(gear, "L21"));
  if (
    !Number.isFinite(tipbackUpper) ||
    !Number.isFinite(tipbackLower) ||
    tipbackUpper >= tipbackLower - ANGLE_TOL
  ) {
    feedback.push(format(STRINGS.gear.tipback, tipbackUpper, tipbackLower));
    failures += 1;
  }

  const rolloverUpper = asNumber(getCell(gear, "M20"));
  const rolloverLower = asNumber(getCell(gear, "M21"));
  if (
    !Number.isFinite(rolloverUpper) ||
    !Number.isFinite(rolloverLower) ||
    rolloverUpper >= rolloverLower - ANGLE_TOL
  ) {
    feedback.push(format(STRINGS.gear.rollover, rolloverUpper, rolloverLower));
    failures += 1;
  }

  const rotationSpeed = asNumber(getCell(gear, "N20"));
  if (!Number.isFinite(rotationSpeed) || rotationSpeed >= 200 - SPEED_TOL) {
    feedback.push(format(STRINGS.gear.rotation, rotationSpeed));
    failures += 1;
  }
  if (!Number.isFinite(rotationSpeed) || rotationSpeed >= 200 - SPEED_TOL) {
    // Advisory echo (mirrors GE3 behavior)
    feedback.push(format(STRINGS.gear.takeoffSpeed, rotationSpeed));
  }

  if (failures > 0) {
    const deduction = Math.min(4, failures);
    feedback.push(format(STRINGS.gear.deduction, deduction));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
