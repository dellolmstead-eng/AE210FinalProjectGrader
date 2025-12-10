import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";
import { getPlanformPoint, computeEdgeAngle } from "../geomUtils.js";

const STEALTH_ANGLE_TOL = 5; // degrees
const PCS_DIHEDRAL_THRESHOLD = 5; // degrees
const VT_TILT_THRESHOLD = 85; // degrees
const EPS = 1e-6;

const normalizeAngle = (angle) => {
  if (!Number.isFinite(angle)) {
    return Number.NaN;
  }
  let normalized = angle % 180;
  if (normalized < 0) {
    normalized += 180;
  }
  return normalized;
};

const areParallel = (angle, wingAngle) => {
  if (!Number.isFinite(angle) || !Number.isFinite(wingAngle)) {
    return false;
  }
  const a = normalizeAngle(angle);
  const b = normalizeAngle(wingAngle);
  const diff = Math.abs(a - b);
  const alt = 180 - diff;
  return Math.min(diff, alt) <= STEALTH_ANGLE_TOL;
};

const pointDifference = (a, b) => ({
  x: b.x - a.x,
  y: b.y - a.y,
});

const normalHitsCenterline = (tip, inner) => {
  if (!Number.isFinite(tip.x) || !Number.isFinite(tip.y) || !Number.isFinite(inner.x) || !Number.isFinite(inner.y)) {
    return false;
  }
  const direction = pointDifference(tip, inner);
  const normals = [
    { x: direction.y, y: -direction.x },
    { x: -direction.y, y: direction.x },
  ];
  return normals.some((normal) => {
    if (Math.abs(normal.y) < EPS) {
      return false;
    }
    const t = -tip.y / normal.y;
    if (t <= 0) {
      return false;
    }
    const x = tip.x + normal.x * t;
    return Number.isFinite(x);
  });
};

export function runStealthChecks(workbook) {
  const main = workbook.sheets.main;
  const geom = workbook.sheets.geom;
  const feedback = [];
  let failures = 0;

  const wingArea = asNumber(getCell(main, "B18"));
  const pcsArea = asNumber(getCell(main, "C18"));
  const strakeArea = asNumber(getCell(main, "D18"));
  const vtArea = asNumber(getCell(main, "H18"));
  const pcsActive = !Number.isFinite(pcsArea) || pcsArea >= 1;
  const strakeActive = !Number.isFinite(strakeArea) || strakeArea >= 1;
  const vtActive = !Number.isFinite(vtArea) || vtArea >= 1;
  const wingActive = !Number.isFinite(wingArea) || wingArea >= 1;

  const wingLeadingAngle = computeEdgeAngle(geom, 38, 39);
  const wingTipTE = getPlanformPoint(geom, 40);
  const wingCenterTE = getPlanformPoint(geom, 41);
  const wingTrailingAngle = computeEdgeAngle(geom, 40, 41);

  const pcsLeadingAngle = computeEdgeAngle(geom, 115, 116);
  const pcsTrailingAngle = computeEdgeAngle(geom, 117, 118);
  const pcsDihedral = asNumber(getCell(main, "C26"));

  const strakeLeadingAngle = computeEdgeAngle(geom, 152, 153);
  const strakeTrailingAngle = computeEdgeAngle(geom, 154, 155);

  const vtLeadingAngle = computeEdgeAngle(geom, 163, 164);
  const vtTrailingAngle = computeEdgeAngle(geom, 165, 166);
  const vtTilt = asNumber(getCell(main, "H27"));

  const recordFailure = (message) => {
    feedback.push(message);
    failures += 1;
  };

  if (pcsActive && wingActive && !areParallel(pcsLeadingAngle, wingLeadingAngle)) {
    recordFailure(format(STRINGS.stealth.pcsSweep, pcsLeadingAngle, wingLeadingAngle, STEALTH_ANGLE_TOL));
  }

  const wingShielded =
    wingActive &&
    (areParallel(wingTrailingAngle, wingLeadingAngle) || normalHitsCenterline(wingTipTE, wingCenterTE));
  if (!wingShielded) {
    recordFailure(format(STRINGS.stealth.wingTrailing, wingTrailingAngle, STEALTH_ANGLE_TOL));
  }

  const checkParallelPair = (angle, template) => {
    if (!Number.isFinite(angle) || !Number.isFinite(wingLeadingAngle)) {
      recordFailure(STRINGS.stealth.missingGeom);
    } else if (!areParallel(angle, wingLeadingAngle)) {
      recordFailure(format(template, angle, wingLeadingAngle, STEALTH_ANGLE_TOL));
    }
  };

  if (pcsActive && Number.isFinite(pcsDihedral) && pcsDihedral > PCS_DIHEDRAL_THRESHOLD) {
    checkParallelPair(pcsLeadingAngle, STRINGS.stealth.pcsLeadingParallel);
    checkParallelPair(pcsTrailingAngle, STRINGS.stealth.pcsTrailingParallel);
  }

  if (strakeActive) {
    checkParallelPair(strakeLeadingAngle, STRINGS.stealth.strakeLeadingParallel);
    checkParallelPair(strakeTrailingAngle, STRINGS.stealth.strakeTrailingParallel);
  }

  if (!vtActive) {
    // ignore
  } else if (!Number.isFinite(vtTilt)) {
    recordFailure(STRINGS.stealth.missingGeom);
  } else if (vtTilt < VT_TILT_THRESHOLD) {
    checkParallelPair(vtLeadingAngle, STRINGS.stealth.vtLeadingParallel);
    checkParallelPair(vtTrailingAngle, STRINGS.stealth.vtTrailingParallel);
  }

  const deduction = Math.min(5, failures);
  if (deduction > 0) {
    feedback.push(format(STRINGS.stealth.deduction, deduction));
  }

  // Stealth feedback is advisory; geometry deductions are handled in attachments.
  return { delta: 0, feedback };
}
