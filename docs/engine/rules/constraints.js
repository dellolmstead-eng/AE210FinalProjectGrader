import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const MACH_TOL = 1e-2;
const VALUE_TOL = 1e-3;
const PS_TOL = 1;
const CDX_TOL = 1e-3;
const PAYLOAD_INT_TOL = 0.01;
const RADIUS_TOL = 0.05;
const BETA_DEFAULT = 0.87620980519917;
const CURVE_LABEL_MAP = {
  MaxMach: "MaxMach",
  CruiseMach: "Supercruise",
  "Cmbt Turn1": "CombatTurn1",
  "Cmbt Turn2": "CombatTurn2",
  Ps1: "Ps1",
  Ps2: "Ps2",
  Takeoff: "Takeoff",
};

const CONSTRAINTS = [
  {
    label: "MaxMach",
    row: 3,
    altMin: 35000,
    machMin: 2.0,
    machObj: 2.2,
    nEq: 1,
    abEq: 100,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "CruiseMach",
    row: 4,
    altMin: 35000,
    machMin: 1.5,
    machObj: 1.8,
    nEq: 1,
    abEq: 0,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Cmbt Turn1",
    row: 6,
    machEq: 1.2,
    altEq: 30000,
    nMin: 3.0,
    nObj: 4.0,
    abEq: 100,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Cmbt Turn2",
    row: 7,
    machEq: 0.9,
    altEq: 10000,
    nMin: 4.0,
    nObj: 4.5,
    abEq: 100,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Ps1",
    row: 8,
    machEq: 1.15,
    altEq: 30000,
    nEq: 1,
    abEq: 100,
    psMin: 400,
    psObj: 500,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Ps2",
    row: 9,
    machEq: 0.9,
    altEq: 10000,
    nEq: 1,
    abEq: 0,
    psMin: 400,
    psObj: 500,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Takeoff",
    row: 12,
    altEq: 0,
    machEq: 1.2,
    nEq: 0.03,
    abEq: 100,
    cdxAllowed: [0, 0.035],
    betaEq: 1,
  },
  {
    label: "Landing",
    row: 13,
    altEq: 0,
    machEq: 1.3,
    nEq: 0.5,
    abEq: 0,
    cdxAllowed: [0, 0.045],
    betaEq: 1,
  },
];

const CURVE_ROWS = [
  { row: 23, label: "MaxMach" },
  { row: 24, label: "Supercruise" },
  { row: 26, label: "CombatTurn1" },
  { row: 27, label: "CombatTurn2" },
  { row: 28, label: "Ps1" },
  { row: 29, label: "Ps2" },
  { row: 32, label: "Takeoff" },
];

function interpolate(xList, yList, x) {
  const pairs = xList
    .map((value, idx) => ({ x: asNumber(value), y: asNumber(yList[idx]) }))
    .filter((pair) => Number.isFinite(pair.x) && Number.isFinite(pair.y))
    .sort((a, b) => a.x - b.x);

  if (pairs.length === 0) {
    return null;
  }
  if (x <= pairs[0].x) {
    if (pairs.length === 1) {
      return pairs[0].y;
    }
    const [p0, p1] = pairs;
    const slope = (p1.y - p0.y) / (p1.x - p0.x);
    return p0.y + slope * (x - p0.x);
  }
  if (x >= pairs[pairs.length - 1].x) {
    const [p0, p1] = pairs.slice(-2);
    const slope = (p1.y - p0.y) / (p1.x - p0.x);
    return p1.y + slope * (x - p1.x);
  }
  for (let i = 0; i < pairs.length - 1; i += 1) {
    const p0 = pairs[i];
    const p1 = pairs[i + 1];
    if (x >= p0.x && x <= p1.x) {
      const slope = (p1.y - p0.y) / (p1.x - p0.x);
      return p0.y + slope * (x - p0.x);
    }
  }
  return null;
}

export function runConstraintChecks(workbook) {
  const feedback = [];
  let tableErrors = 0;
  let payloadPenalty = 0;
  let curvePenalty = 0;
  const rowErrors = {};
  const objectiveLabels = new Set();
  const curveStatus = {};

  const main = workbook.sheets.main;
  const consts = workbook.sheets.consts;
  const fuelAvailable = asNumber(getCell(main, "O18"));
  const fuelCapacity = asNumber(getCell(main, "O15"));
  const betaDefault =
    Number.isFinite(fuelAvailable) && Number.isFinite(fuelCapacity) && fuelCapacity !== 0
      ? 1 - fuelAvailable / (2 * fuelCapacity)
      : BETA_DEFAULT;

  const radius = asNumber(getCell(main, "Y37"));
  if (Number.isFinite(radius)) {
    if (radius < 375 - RADIUS_TOL) {
      feedback.push(format(STRINGS.constraint.radiusLow, radius));
      tableErrors += 1;
    } else if (radius >= 410 - RADIUS_TOL) {
      feedback.push(format(STRINGS.constraint.radiusObj, radius));
    }
  }

  const aim120Raw = asNumber(getCell(main, "AB3"));
  const aim9Raw = asNumber(getCell(main, "AB4"));
  const aim120Int =
    Number.isFinite(aim120Raw) && Math.abs(aim120Raw - Math.round(aim120Raw)) <= PAYLOAD_INT_TOL
      ? Math.round(aim120Raw)
      : null;
  const aim9Int =
    Number.isFinite(aim9Raw) && Math.abs(aim9Raw - Math.round(aim9Raw)) <= PAYLOAD_INT_TOL
      ? Math.round(aim9Raw)
      : null;
  if (aim120Int == null || aim9Int == null) {
    const invalidMsg =
      STRINGS.constraint.payloadPenaltyInvalid || "Payload counts for AIM-120 and AIM-9 must be integers.";
    feedback.push(invalidMsg);
    payloadPenalty -= 4;
  } else if (aim120Int < 8) {
    payloadPenalty -= 4;
    feedback.push(format(STRINGS.constraint.payloadPenalty, aim120Int));
  } else if (aim9Int >= 2) {
    feedback.push(format(STRINGS.constraint.payloadObj, aim120Int, aim9Int));
  }

  const distanceObjectives = [];
  const takeoffDist = asNumber(getCell(main, "X12"));
  if (Number.isFinite(takeoffDist)) {
    if (takeoffDist > 3000 + RADIUS_TOL) {
      feedback.push(format(STRINGS.constraint.takeoffHigh, takeoffDist));
      tableErrors += 1;
    } else if (takeoffDist <= 2500 + RADIUS_TOL) {
      distanceObjectives.push({ label: "Takeoff", text: format(STRINGS.constraint.takeoffObj, takeoffDist) });
    }
  }

  const landingDist = asNumber(getCell(main, "X13"));
  if (Number.isFinite(landingDist)) {
    if (landingDist > 5000 + RADIUS_TOL) {
      feedback.push(format(STRINGS.constraint.landingHigh, landingDist));
      tableErrors += 1;
    } else if (landingDist <= 3500 + RADIUS_TOL) {
      distanceObjectives.push({ label: "Landing", text: format(STRINGS.constraint.landingObj, landingDist) });
    }
  }

  CONSTRAINTS.forEach((constraint) => {
    const row = constraint.row;
    const label = constraint.label;
    const beta = asNumber(getCellByIndex(main, row, 19));
    const altitude = asNumber(getCellByIndex(main, row, 20));
    const mach = asNumber(getCellByIndex(main, row, 21));
    const n = asNumber(getCellByIndex(main, row, 22));
    const ab = asNumber(getCellByIndex(main, row, 23));
    const ps = asNumber(getCellByIndex(main, row, 24));
    const cdx = asNumber(getCellByIndex(main, row, 25));

    if (!(label in rowErrors)) {
      rowErrors[label] = false;
    }

    const markError = (message) => {
      feedback.push(message);
      tableErrors += 1;
      rowErrors[label] = true;
    };

    const exceedsTol = (value, expected, tol) =>
      !Number.isFinite(value) || Math.abs(value - expected) > tol;

    if (constraint.machEq != null) {
      if (exceedsTol(mach, constraint.machEq, MACH_TOL)) {
        markError(format(STRINGS.constraint.machEq, label, mach ?? NaN, constraint.machEq));
      }
    } else if (constraint.machMin != null) {
      if (!Number.isFinite(mach) || mach < constraint.machMin - MACH_TOL) {
        markError(format(STRINGS.constraint.machMin, label, mach ?? NaN, constraint.machMin));
      } else if (constraint.machObj != null && Number.isFinite(mach) && mach >= constraint.machObj - MACH_TOL) {
        objectiveLabels.add(label);
      }
    }

    if (constraint.altEq != null) {
      if (exceedsTol(altitude, constraint.altEq, 1)) {
        markError(format(STRINGS.constraint.altEq, label, altitude ?? NaN, constraint.altEq));
      }
    } else if (constraint.altMin != null) {
      if (!Number.isFinite(altitude) || altitude < constraint.altMin - 1) {
        markError(format(STRINGS.constraint.altMin, label, altitude ?? NaN, constraint.altMin));
      }
    }

    if (constraint.nEq != null) {
      if (exceedsTol(n, constraint.nEq, VALUE_TOL)) {
        markError(format(STRINGS.constraint.nEq, label, n ?? NaN, constraint.nEq));
      }
    } else if (constraint.nMin != null) {
      if (!Number.isFinite(n) || n < constraint.nMin - VALUE_TOL) {
        markError(format(STRINGS.constraint.nMin, label, n ?? NaN, constraint.nMin));
      } else if (constraint.nObj != null && Number.isFinite(n) && n >= constraint.nObj - VALUE_TOL) {
        objectiveLabels.add(label);
      }
    }

    if (constraint.abEq != null) {
      if (exceedsTol(ab, constraint.abEq, VALUE_TOL)) {
        markError(format(STRINGS.constraint.abEq, label, ab ?? NaN, constraint.abEq));
      }
    }

    if (constraint.psEq != null) {
      if (exceedsTol(ps, constraint.psEq, PS_TOL)) {
        markError(format(STRINGS.constraint.psEq, label, ps ?? NaN, constraint.psEq));
      }
    } else if (constraint.psMin != null) {
      if (!Number.isFinite(ps) || ps < constraint.psMin - PS_TOL) {
        markError(format(STRINGS.constraint.psMin, label, ps ?? NaN, constraint.psMin));
      } else if (constraint.psObj != null && Number.isFinite(ps) && ps >= constraint.psObj - PS_TOL) {
        objectiveLabels.add(label);
      }
    }

    if (constraint.betaEq != null) {
      const betaTarget = constraint.betaEq === BETA_DEFAULT ? betaDefault : constraint.betaEq;
      if (exceedsTol(beta, betaTarget, VALUE_TOL)) {
        markError(format(STRINGS.constraint.betaEq, label, betaTarget ?? NaN, beta ?? NaN));
      }
    }

    if (constraint.cdxAllowed) {
      const allowedMatch = constraint.cdxAllowed.some(
        (allowed) => Number.isFinite(cdx) && Math.abs(cdx - allowed) <= CDX_TOL
      );
      if (!allowedMatch) {
        const allowedText = constraint.cdxAllowed.map((value) => value.toFixed(3)).join(", ");
        markError(format(STRINGS.constraint.cdxAllowed, label, cdx ?? NaN, allowedText));
      }
    } else if (constraint.cdxEq != null) {
      if (exceedsTol(cdx, constraint.cdxEq, CDX_TOL)) {
        markError(format(STRINGS.constraint.cdxEq, label, cdx ?? NaN, constraint.cdxEq));
      }
    }
  });

  const deduction = Math.min(2, tableErrors);
  let delta = 0;
  if (deduction > 0) {
    feedback.push(format(STRINGS.constraintSummary, deduction));
    delta -= deduction;
  }

  const failures = [];
  try {
    const wsAxis = [];
    for (let col = 11; col <= 31; col += 1) {
      wsAxis.push(getCellByIndex(consts, 22, col));
    }
    const wsDesign = asNumber(getCell(main, "P13"));
    const twDesign = asNumber(getCell(main, "Q13"));

    if (Number.isFinite(wsDesign) && Number.isFinite(twDesign)) {
      CURVE_ROWS.forEach(({ row, label }) => {
        const twCurve = [];
        for (let col = 11; col <= 31; col += 1) {
          twCurve.push(getCellByIndex(consts, row, col));
        }
        const requiredTW = interpolate(wsAxis, twCurve, wsDesign);
        if (requiredTW != null) {
          const passes = twDesign >= requiredTW - VALUE_TOL;
          curveStatus[label] = passes;
          if (!passes) {
            failures.push(label);
          }
        }
      });

      const wsLimitLanding = asNumber(getCell(consts, "L33"));
      if (Number.isFinite(wsLimitLanding)) {
        const landingPass = wsDesign <= wsLimitLanding + VALUE_TOL;
        curveStatus.Landing = landingPass;
        if (!landingPass) {
          failures.push("Landing");
          feedback.push(format(STRINGS.constraint.landingCurve, wsDesign, wsLimitLanding));
        }
      }

      if (failures.length === 1) {
        const deduction = 4;
        curvePenalty -= deduction;
        feedback.push(
          format(STRINGS.constraint.curveFailure, deduction, "", failures[0]) + STRINGS.constraint.curveSuffixFew
        );
      } else if (failures.length >= 2) {
        const deduction = 8;
        curvePenalty -= deduction;
        const joined = failures.join(", ");
        const suffix =
          failures.length > 6 ? STRINGS.constraint.curveSuffixMany : STRINGS.constraint.curveSuffixFew;
        feedback.push(format(STRINGS.constraint.curveFailure, deduction, "s", joined) + suffix);
      }
    }
  } catch (err) {
    feedback.push(format(STRINGS.constraint.curveError, err.message));
  }

  // Distance objectives only if the corresponding curve passed
  distanceObjectives.forEach(({ label, text }) => {
    const curveLabel = CURVE_LABEL_MAP[label] || label;
    if (curveStatus[curveLabel] !== false) {
      feedback.push(text);
    }
  });

  objectiveLabels.forEach((label) => {
    const curveLabel = CURVE_LABEL_MAP[label] || label;
    const status = curveStatus[curveLabel];
    if (status === true && rowErrors[label] !== true) {
      feedback.push(format(STRINGS.constraint.objectiveBonus, label));
    } else if (status === false) {
      feedback.push(format(STRINGS.constraint.objectiveFail, label));
    }
  });

  return {
    delta,
    payloadDelta: payloadPenalty,
    curveDelta: curvePenalty,
    feedback,
  };
}
