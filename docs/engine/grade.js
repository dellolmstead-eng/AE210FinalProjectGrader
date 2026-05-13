import { getCell, asNumber } from "./parseUtils.js";
import { pchip } from "./pchip.js";
import { runStealthChecks } from "./rules/stealth.js";

// Utility: convert zero-based row/col to Excel ref
function cellRef(rowIdx, colIdx) {
  let col = "";
  let n = colIdx;
  while (n >= 0) {
    col = String.fromCharCode((n % 26) + 65) + col;
    n = Math.floor(n / 26) - 1;
  }
  return `${col}${rowIdx + 1}`;
}

const TOL = {
  eq: 1e-3,
  wto: 1e-3,
  alt: 1,
  mach: 1e-2,
  time: 1e-2,
  dist: 1e-3,
  fuel: 5e-2,
};

const BETA_DEFAULT = 0.87620980519917;
const BASE_TOTAL = 45;
const OBJECTIVE_TOTAL = 11;
const NON_VIABLE_CAP = BASE_TOTAL / 2;

const roundToTenth = (value) => (Number.isFinite(value) ? Math.round(value * 10) / 10 : 0);
const ternary = (cond, a, b) => (cond ? a : b);
const clamp01 = (value) => Math.max(0, Math.min(1, value));
const formatScore = (value) => (Math.abs(value - Math.round(value)) < 1e-9 ? `${Math.round(value)}` : value.toFixed(1));
const fmt1 = (value) => (Number.isFinite(value) ? value.toFixed(1) : `${roundToTenth(value)}`);
const fmt2 = (value) => (Number.isFinite(value) ? value.toFixed(2) : `${roundToTenth(value)}`);
const matlabFixed = (value, digits) => {
  if (!Number.isFinite(value)) return `${roundToTenth(value)}`;
  const scale = 10 ** digits;
  const scaled = value * scale;
  const floor = Math.floor(scaled);
  const frac = scaled - floor;
  if (Math.abs(frac - 0.5) < 1e-9) {
    const rounded = floor % 2 === 0 ? floor : floor + 1;
    return (rounded / scale).toFixed(digits);
  }
  return value.toFixed(digits);
};

const getNumber = (sheet, ref) => asNumber(getCell(sheet, ref));

function linearBonus(value, threshold, objective) {
  if (!Number.isFinite(value)) return 0;
  if (Math.abs(objective - threshold) < Number.EPSILON) return value >= objective ? 1 : 0;
  return clamp01((value - threshold) / (objective - threshold));
}

function linearBonusInv(value, threshold, objective) {
  if (!Number.isFinite(value)) return 0;
  if (Math.abs(objective - threshold) < Number.EPSILON) return value <= objective ? 1 : 0;
  return clamp01((threshold - value) / (threshold - objective));
}

const checkInvalidMainCells = (mainSheet) => {
  const invalidCells = [];
  mainSheet?.forEach((row, rIdx) => {
    if (rIdx > 74) return; // Only meaningful Main-sheet content is within A1:AG75.
    if (!row) return;
    row.forEach((value, cIdx) => {
      if (cIdx > 32) return;
      if (typeof value === "string" && /^#(DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|N\/A)$/i.test(value.trim())) {
        invalidCells.push(cellRef(rIdx, cIdx));
      } else if (typeof value === "number" && !Number.isFinite(value)) {
        invalidCells.push(cellRef(rIdx, cIdx));
      }
    });
  });
  return invalidCells;
};

function checkGeometryBlocks(main) {
  const missing1 = [];
  const missing2 = [];
  const pushMissing1 = (row, col) => missing1.push(cellRef(row - 1, col - 1));
  const pushMissing2 = (row, col) => missing2.push(cellRef(row - 1, col - 1));

  // Block B18:H27 with skips
  const skips1 = new Set(["B24", "C24", "E22", "G22", "D27", "E27", "F27", "G27", "H26"]);
  for (let r = 18; r <= 27; r += 1) {
    for (let c = 2; c <= 8; c += 1) {
      const ref = `${String.fromCharCode(64 + c)}${r}`;
      if (skips1.has(ref)) continue;
      if (!Number.isFinite(getNumber(main, ref))) pushMissing1(r, c);
    }
  }

  // Block C34:F53
  for (let r = 34; r <= 53; r += 1) {
    for (let c = 3; c <= 6; c += 1) {
      const ref = `${String.fromCharCode(64 + c)}${r}`;
      if (!Number.isFinite(getNumber(main, ref))) pushMissing2(r, c);
    }
  }

  return { missing1, missing2 };
}

function checkMissionProfile(main, radius, betaExpected) {
  const feedback = [];
  const LEG_COLUMNS = [11, 12, 13, 14, 16, 18, 19, 22, 23]; // 1-based Main-sheet columns used by the GE mission table
  const readRowValues = (row1) => LEG_COLUMNS.map((col1) => asNumber(main?.[row1 - 1]?.[col1 - 1]));

  const alt = readRowValues(33);
  const mach = readRowValues(35);
  const ab = readRowValues(36);
  const dist = readRowValues(38);
  const time = readRowValues(39);
  const combatTurnAngle = asNumber(main?.[38]?.[27]); // AB39
  const supercruiseMach = asNumber(main?.[3]?.[20]); // U4

  let missionErrors = 0;
  if (Math.abs(alt[0]) > TOL.alt || Math.abs(ab[0] - 100) > TOL.eq) {
    feedback.push(`Leg 1: Altitude must be 0 ft and AB = 100% (found alt=${fmt1(alt[0])}, AB=${fmt1(ab[0])})`);
    missionErrors += 1;
  }

  if (alt[1] < Math.min(alt[0], alt[2]) - TOL.alt || alt[1] > Math.max(alt[0], alt[2]) + TOL.alt) {
    feedback.push(`Leg 2: Altitude must remain between legs 1 and 3 (found alt2=${fmt1(alt[1])}, alt1=${fmt1(alt[0])}, alt3=${fmt1(alt[2])})`);
    missionErrors += 1;
  }
  if (mach[1] < Math.min(mach[0], mach[2]) - TOL.mach || mach[1] > Math.max(mach[0], mach[2]) + TOL.mach) {
    feedback.push(`Leg 2: Mach must remain between legs 1 and 3 (found mach2=${fmt2(mach[1])}, mach1=${fmt2(mach[0])}, mach3=${fmt2(mach[2])})`);
    missionErrors += 1;
  }
  if (Math.abs(ab[1]) > TOL.eq) {
    feedback.push(`Leg 2: AB must be 0% (found AB=${fmt1(ab[1])})`);
    missionErrors += 1;
  }

  if (alt[2] < 35000 - TOL.alt || Math.abs(mach[2] - 0.9) > TOL.mach || Math.abs(ab[2]) > TOL.eq) {
    feedback.push(`Leg 3: Must be >= 35,000 ft, Mach = 0.9, AB = 0% (found alt=${fmt1(alt[2])}, mach=${fmt2(mach[2])}, AB=${fmt1(ab[2])})`);
    missionErrors += 1;
  }

  if (alt[3] < 35000 - TOL.alt || Math.abs(mach[3] - 0.9) > TOL.mach || Math.abs(ab[3]) > TOL.eq) {
    feedback.push(`Leg 4: Must be >= 35,000 ft, Mach = 0.9, AB = 0% (found alt=${fmt1(alt[3])}, mach=${fmt2(mach[3])}, AB=${fmt1(ab[3])})`);
    missionErrors += 1;
  }

  if (alt[4] < 35000 - TOL.alt || !Number.isFinite(supercruiseMach) || Math.abs(mach[4] - supercruiseMach) > TOL.mach || Math.abs(ab[4]) > TOL.eq || dist[4] < 150 - TOL.dist) {
    feedback.push(`Leg 5: Must be >= 35,000 ft, Mach = constraint Supercruise Mach (Main!U4), AB = 0%, Distance >= 150 nm (found alt=${fmt1(alt[4])}, mach=${fmt2(mach[4])}, AB=${fmt1(ab[4])}, dist=${fmt1(dist[4])})`);
    missionErrors += 1;
  }

  if (alt[5] < 30000 - TOL.alt || mach[5] < 1.2 - TOL.mach || Math.abs(ab[5] - 100) > TOL.eq) {
    feedback.push(`Leg 6: Must be >= 30,000 ft, Mach >= 1.2, AB = 100% (found alt=${fmt1(alt[5])}, mach=${fmt2(mach[5])}, AB=${fmt1(ab[5])})`);
    missionErrors += 1;
  }
  if (!(combatTurnAngle >= 720)) {
    feedback.push("Two full 360 turns are required. Increase total turn angle (cell AB39) to 720 degrees or greater to meet the combat turn requirement");
    missionErrors += 1;
  }

  if (alt[6] < 35000 - TOL.alt || !Number.isFinite(supercruiseMach) || Math.abs(mach[6] - supercruiseMach) > TOL.mach || Math.abs(ab[6]) > TOL.eq || dist[6] < 150 - TOL.dist) {
    feedback.push(`Leg 7: Must be >= 35,000 ft, Mach = constraint Supercruise Mach (Main!U4), AB = 0%, Distance >= 150 nm (found alt=${fmt1(alt[6])}, mach=${fmt2(mach[6])}, AB=${fmt1(ab[6])}, dist=${fmt1(dist[6])})`);
    missionErrors += 1;
  }

  if (alt[7] < 35000 - TOL.alt || Math.abs(mach[7] - 0.9) > TOL.mach || Math.abs(ab[7]) > TOL.eq) {
    feedback.push(`Leg 8: Must be >= 35,000 ft, Mach = 0.9, AB = 0% (found alt=${fmt1(alt[7])}, mach=${fmt2(mach[7])}, AB=${fmt1(ab[7])})`);
    missionErrors += 1;
  }

  if (Math.abs(alt[8] - 10000) > TOL.alt || Math.abs(mach[8] - 0.4) > TOL.mach || Math.abs(ab[8]) > TOL.eq || Math.abs(time[8] - 20) > TOL.time) {
    feedback.push(`Leg 9: Must be 10,000 ft, Mach = 0.4, AB = 0%, Time = 20 min (found alt=${fmt1(alt[8])}, mach=${fmt2(mach[8])}, AB=${fmt1(ab[8])}, time=${fmt2(time[8])})`);
    missionErrors += 1;
  }

  let rangePass = false;
  let rangeObjectivePass = false;
  if (Number.isFinite(radius)) {
    if (radius >= 410 - TOL.dist) {
      rangePass = true;
      rangeObjectivePass = true;
      feedback.push(`Mission radius meets objective (410 nm): ${Number(radius).toFixed(1)}`);
    } else if (radius >= 375 - TOL.dist) {
      rangePass = true;
    } else {
      feedback.push(`Mission radius below threshold (375 nm): ${Number(radius).toFixed(1)}`);
      missionErrors += 1;
    }
  } else {
    feedback.push("Mission radius missing; unable to verify range requirement.");
    missionErrors += 1;
  }

  const missionPass = missionErrors === 0;
  return { feedback, missionPass, missionErrors, rangePass, rangeObjectivePass, betaExpected };
}

function checkEfficiency(main) {
  const feedback = [];
  const o1 = getNumber(main, "O1");
  const q1 = getNumber(main, "Q1");
  const c30 = getNumber(main, "C30");
  const d30 = getNumber(main, "D30");
  let pass = true;
  if (!Number.isFinite(o1) || Math.abs(o1 - 0.0037) > TOL.eq) {
    feedback.push(`O1 must be 0.0037 (found ${roundToTenth(o1)})`);
    pass = false;
  }
  if (!Number.isFinite(q1) || Math.abs(q1 - 2.2) > TOL.eq) {
    feedback.push(`Q1 must be 2.2 (found ${roundToTenth(q1)})`);
    pass = false;
  }
  if (!Number.isFinite(c30) || Math.abs(c30 - 0.8) > TOL.eq) {
    feedback.push(`C30 must be 0.8 (found ${roundToTenth(c30)})`);
    pass = false;
  }
  if (!Number.isFinite(d30) || Math.abs(d30 - 2.0) > TOL.eq) {
    feedback.push(`D30 must be 2.0 (found ${roundToTenth(d30)})`);
    pass = false;
  }
  return { pass, feedback };
}

function checkThrust(miss) {
  const feedback = [];
  const thrustShort = [];
  for (let c = 2; c <= 13; c += 1) {
    const drag = asNumber(miss?.[47]?.[c]);
    const available = asNumber(miss?.[48]?.[c]);
    if (!Number.isFinite(available) || !Number.isFinite(drag)) continue;
    if (available <= drag + TOL.eq) thrustShort.push(c);
  }
  if (thrustShort.length > 0) {
    feedback.push(`Thrust shortfall: Tavailable <= Drag for ${thrustShort.length} mission segment(s).`);
  }
  return { pass: thrustShort.length === 0, feedback };
}

function checkAero(aero) {
  let issues = 0;
  if (aero?.[2]?.[6] === aero?.[3]?.[6]) issues += 1; // G3 vs G4
  if (aero?.[9]?.[6] === aero?.[10]?.[6]) issues += 1; // G10 vs G11
  if (aero?.[14]?.[0] === aero?.[15]?.[0]) issues += 1; // A15 vs A16
  const deduction = Math.min(3, issues);
  return {
    deduction,
    feedback: deduction > 0 ? [`-${deduction} pts Aero tab formulas not active in cells A15, G3, or G10`] : [],
  };
}

function checkControlAttachment(main, geom) {
  const fb = [];
  let failures = 0;
  const VALUE_TOL = 1e-3;
  const AR_TOL = 0.1;
  const VT_WING_FRACTION = 0.8;
  const EDGE_ALIGN_TOL = 0.2;

  const planformPoint = (row) => {
    const x = asNumber(geom?.[row - 1]?.[11]);
    const yCandidates = [asNumber(geom?.[row - 1]?.[12]), asNumber(geom?.[row - 1]?.[13])].filter(Number.isFinite);
    return [x, yCandidates.length === 0 ? 0 : Math.max(...yCandidates.map((value) => Math.abs(value)))];
  };
  const sortEdgePairsByY = (relevantA, relevantB, oppositeA, oppositeB) => {
    if (relevantA[1] <= relevantB[1]) {
      return {
        relevantInboard: relevantA,
        relevantOutboard: relevantB,
        oppositeInboard: oppositeA,
        oppositeOutboard: oppositeB,
      };
    }
    return {
      relevantInboard: relevantB,
      relevantOutboard: relevantA,
      oppositeInboard: oppositeB,
      oppositeOutboard: oppositeA,
    };
  };
  const interpolateEdgeXAtY = (pointA, pointB, y) => {
    if (!pointA.every(Number.isFinite) || !pointB.every(Number.isFinite) || !Number.isFinite(y)) {
      return { x: Number.NaN, inRange: false };
    }
    const y1 = pointA[1];
    const y2 = pointB[1];
    const lower = Math.min(y1, y2);
    const upper = Math.max(y1, y2);
    if (y < lower - 1e-6 || y > upper + 1e-6) {
      return { x: Number.NaN, inRange: false };
    }
    if (Math.abs(y2 - y1) < 1e-9) {
      return { x: pointA[0], inRange: Math.abs(y - y1) <= 1e-6 };
    }
    const t = (y - y1) / (y2 - y1);
    return { x: pointA[0] + t * (pointB[0] - pointA[0]), inRange: true };
  };
  const checkWingDevicePlacement = (deviceName, edgeName, relevantA, relevantB, oppositeA, oppositeB, wingLeadingRoot, wingLeadingTip, wingTrailingRoot, wingTrailingTip) => {
    const points = [relevantA, relevantB, oppositeA, oppositeB, wingLeadingRoot, wingLeadingTip, wingTrailingRoot, wingTrailingTip];
    if (points.some((point) => !point.every(Number.isFinite))) {
      return { failed: true, messages: [`Unable to verify ${deviceName} wing placement due to missing geometry data`] };
    }
    const { relevantInboard, relevantOutboard, oppositeInboard, oppositeOutboard } =
      sortEdgePairsByY(relevantA, relevantB, oppositeA, oppositeB);
    const wingSpan = Math.max(wingLeadingRoot[1], wingLeadingTip[1], wingTrailingRoot[1], wingTrailingTip[1]);
    let spanFail = false;
    let edgeFail = false;
    let envelopeFail = false;
    for (const [relevantPoint, oppositePoint] of [
      [relevantInboard, oppositeInboard],
      [relevantOutboard, oppositeOutboard],
    ]) {
      const y = relevantPoint[1];
      if (y < -EDGE_ALIGN_TOL || y > wingSpan + EDGE_ALIGN_TOL) {
        spanFail = true;
        continue;
      }
      const wingLeading = interpolateEdgeXAtY(wingLeadingRoot, wingLeadingTip, y);
      const wingTrailing = interpolateEdgeXAtY(wingTrailingRoot, wingTrailingTip, y);
      if (!wingLeading.inRange || !wingTrailing.inRange) {
        spanFail = true;
        continue;
      }
      const targetX = edgeName === "leading edge" ? wingLeading.x : wingTrailing.x;
      if (Math.abs(relevantPoint[0] - targetX) > EDGE_ALIGN_TOL) {
        edgeFail = true;
      }
      const lower = Math.min(wingLeading.x, wingTrailing.x) - EDGE_ALIGN_TOL;
      const upper = Math.max(wingLeading.x, wingTrailing.x) + EDGE_ALIGN_TOL;
      if (oppositePoint[0] < lower || oppositePoint[0] > upper) {
        envelopeFail = true;
      }
    }
    const messages = [];
    if (spanFail) {
      messages.push(`${deviceName} extends outside the wing span in top view.`);
    }
    if (edgeFail) {
      messages.push(`${deviceName} ${edgeName} must align with the wing ${edgeName} within ${EDGE_ALIGN_TOL.toFixed(2)} ft in top view.`);
    }
    if (envelopeFail) {
      messages.push(`${deviceName} must remain within the wing planform envelope in top view.`);
    }
    return { failed: spanFail || edgeFail || envelopeFail, messages };
  };

  const fuselage_end = getNumber(main, "B32");
  const pcsArea = getNumber(main, "C18");
  const vtArea = getNumber(main, "H18");
  const strakeArea = getNumber(main, "D18");
  const PCS_x = getNumber(main, "C23");
  const PCS_root = getNumber(geom, "C8");
  if (Number.isFinite(pcsArea) && pcsArea >= 1) {
  if (!Number.isFinite(fuselage_end) || !Number.isFinite(PCS_x) || !Number.isFinite(PCS_root)) {
    fb.push("Unable to verify PCS placement due to missing geometry data");
    failures += 1;
  } else if (PCS_x > fuselage_end - 0.25 * PCS_root) {
    fb.push("PCS X-location too far aft. Must overlap at least 25% of root chord.");
    failures += 1;
  }
  }

  const VT_x = getNumber(main, "H23");
  const VT_root = getNumber(geom, "C10");
  if (Number.isFinite(vtArea) && vtArea >= 1) {
  if (!Number.isFinite(fuselage_end) || !Number.isFinite(VT_x) || !Number.isFinite(VT_root)) {
    fb.push("Unable to verify vertical tail placement due to missing geometry data");
    failures += 1;
  } else if (VT_x > fuselage_end - 0.25 * VT_root) {
    fb.push("VT X-location too far aft. Must overlap at least 25% of root chord.");
    failures += 1;
  }
  }

  const PCS_z = getNumber(main, "C25");
  const fuse_z_center = getNumber(main, "D52");
  const fuse_z_height = getNumber(main, "F52");
  if (Number.isFinite(pcsArea) && pcsArea >= 1) {
  if (!Number.isFinite(PCS_z) || !Number.isFinite(fuse_z_center) || !Number.isFinite(fuse_z_height)) {
    fb.push("Unable to verify PCS vertical placement due to missing geometry data");
    failures += 1;
  } else if (PCS_z < fuse_z_center - fuse_z_height / 2 || PCS_z > fuse_z_center + fuse_z_height / 2) {
    fb.push("PCS Z-location outside fuselage vertical bounds.");
    failures += 1;
  }
  }

  const VT_y = getNumber(main, "H24");
  const fuse_width = getNumber(main, "E52");
  let vtMountedOffFuselage = false;
  if (Number.isFinite(vtArea) && vtArea >= 1) {
  if (!Number.isFinite(VT_y) || !Number.isFinite(fuse_width)) {
    fb.push("Unable to verify vertical tail lateral placement due to missing geometry data");
    failures += 1;
  } else if (Math.abs(VT_y) > fuse_width / 2 + VALUE_TOL) {
    vtMountedOffFuselage = true;
    fb.push("Vertical tail mounted off the fuselage; ensure structural support at the wing.");
  }
  }

  if (Number.isFinite(strakeArea) && strakeArea >= 1) {
    const sweep = getNumber(geom, "K15");
    const y = getNumber(geom, "M152");
    const strake = getNumber(geom, "L155");
    const apex = getNumber(geom, "L38");
    if (!Number.isFinite(sweep) || !Number.isFinite(y) || !Number.isFinite(strake) || !Number.isFinite(apex)) {
      fb.push("Unable to verify strake attachment due to missing geometry data");
      failures += 1;
    } else {
      const wing = y / Math.tan(((90 - sweep) * Math.PI) / 180) + apex;
      if (wing >= strake + 0.5) {
        fb.push("Strake disconnected.");
        failures += 1;
      }
    }
  }

  const wingLeadingRoot = planformPoint(38);
  const wingLeadingTip = planformPoint(39);
  const wingTrailingRoot = planformPoint(41);
  const wingTrailingTip = planformPoint(40);

  const elevonArea = getNumber(main, "E18");
  if (Number.isFinite(elevonArea) && elevonArea >= 1) {
    const result = checkWingDevicePlacement(
      "Elevon",
      "trailing edge",
      planformPoint(177),
      planformPoint(176),
      planformPoint(174),
      planformPoint(175),
      wingLeadingRoot,
      wingLeadingTip,
      wingTrailingRoot,
      wingTrailingTip,
    );
    fb.push(...result.messages);
    failures += result.failed ? 1 : 0;
  }

  const lefArea = getNumber(main, "F18");
  if (Number.isFinite(lefArea) && lefArea >= 1) {
    const result = checkWingDevicePlacement(
      "LE Flap",
      "leading edge",
      planformPoint(186),
      planformPoint(187),
      planformPoint(189),
      planformPoint(188),
      wingLeadingRoot,
      wingLeadingTip,
      wingTrailingRoot,
      wingTrailingTip,
    );
    fb.push(...result.messages);
    failures += result.failed ? 1 : 0;
  }

  const tefArea = getNumber(main, "G18");
  if (Number.isFinite(tefArea) && tefArea >= 1) {
    const result = checkWingDevicePlacement(
      "TE Flap",
      "trailing edge",
      planformPoint(201),
      planformPoint(200),
      planformPoint(198),
      planformPoint(199),
      wingLeadingRoot,
      wingLeadingTip,
      wingTrailingRoot,
      wingTrailingTip,
    );
    fb.push(...result.messages);
    failures += result.failed ? 1 : 0;
  }

  const component_positions = main?.[22]?.slice(1, 8).map(asNumber) || [];
  const component_areas = main?.[17]?.slice(1, 8).map(asNumber) || [];
  const active_positions = component_positions.filter((_, idx) => Number.isFinite(component_areas[idx]) && component_areas[idx] >= 1);
  if (active_positions.length > 0 && !Number.isFinite(fuselage_end)) {
    fb.push("Unable to verify component X-location due to missing fuselage length");
    failures += 1;
  } else if (active_positions.some((v) => Number.isFinite(v) && v >= fuselage_end)) {
      fb.push(`One or more components X-location extend beyond the fuselage end (B32 = ${Number.isFinite(fuselage_end) ? fuselage_end.toFixed(2) : roundToTenth(fuselage_end)})`);
    failures += 1;
  }

  if (vtMountedOffFuselage) {
    const vtApex = [asNumber(geom?.[162]?.[11]), asNumber(geom?.[162]?.[12])];
    const vtRootTE = [asNumber(geom?.[165]?.[11]), asNumber(geom?.[165]?.[12])];
    const wingTE = [asNumber(geom?.[40]?.[11]), asNumber(geom?.[40]?.[12])];
    if (vtApex.some(Number.isNaN) || vtRootTE.some(Number.isNaN) || wingTE.some(Number.isNaN)) {
      fb.push("Unable to verify vertical tail overlap with wing due to missing geometry data");
      failures += 1;
    } else {
      const chord = vtRootTE[0] - vtApex[0];
      const overlap = Math.max(0, Math.min(wingTE[0], vtRootTE[0]) - vtApex[0]);
      if (!(chord > 0) || overlap + VALUE_TOL < VT_WING_FRACTION * chord) {
        fb.push("Vertical tail mounted on the wing must overlap at least 80% of its root chord with the wing trailing edge.");
        failures += 1;
      }
    }
  }

  const wingAR = getNumber(main, "B19");
  const pcsAR = getNumber(main, "C19");
  const vtAR = getNumber(main, "H19");
  if (Number.isFinite(wingAR) && Number.isFinite(pcsAR) && pcsAR > wingAR + AR_TOL) {
    fb.push(`Pitch control surface aspect ratio (${pcsAR.toFixed(2)}) must be lower than wing aspect ratio (${wingAR.toFixed(2)}).`);
    failures += 1;
  }
  if (Number.isFinite(wingAR) && Number.isFinite(vtAR) && vtAR >= wingAR - AR_TOL) {
    fb.push(`Vertical tail aspect ratio (${vtAR.toFixed(2)}) must be lower than wing aspect ratio (${wingAR.toFixed(2)}).`);
    failures += 1;
  }

  const engine_diameter = getNumber(main, "H29");
  const inlet_x = getNumber(main, "F31");
  const compressor_x = getNumber(main, "F32");
  const engine_start = inlet_x + compressor_x;
  const widths = [];
  for (let r = 34; r <= 53; r += 1) {
    const station_x = asNumber(main?.[r - 1]?.[1]);
    const width = asNumber(main?.[r - 1]?.[4]);
    if (Number.isFinite(station_x) && Number.isFinite(width) && station_x >= engine_start) widths.push(width);
  }
  if (widths.length === 0 || !Number.isFinite(engine_diameter)) {
    fb.push("Unable to verify fuselage width clearance for engines");
    failures += 1;
  } else {
    const minWidth = Math.min(...widths);
    const maxWidth = Math.max(...widths);
    const requiredWidth = engine_diameter + 0.5;
    if (minWidth + VALUE_TOL <= requiredWidth) {
      fb.push(`Fuselage minimum width (${minWidth.toFixed(2)} ft) must exceed engine diameter + 0.5 ft (${requiredWidth.toFixed(2)} ft).`);
      failures += 1;
    }
    const allowedOverhang = 2 * maxWidth;
    if (Number.isFinite(fuselage_end)) {
      const pcsTipX = Math.max(asNumber(geom?.[116]?.[11]), asNumber(geom?.[117]?.[11]));
      const vtTipX = Math.max(asNumber(geom?.[164]?.[11]), asNumber(geom?.[165]?.[11]));
      if (Number.isFinite(pcsTipX)) {
        const overhang = pcsTipX - fuselage_end;
        if (overhang > allowedOverhang + VALUE_TOL) {
          fb.push(`Pitch control surface extends ${overhang.toFixed(2)} ft beyond the fuselage end (limit ${allowedOverhang.toFixed(2)} ft).`);
          failures += 1;
        }
      }
      if (Number.isFinite(vtTipX)) {
        const overhang = vtTipX - fuselage_end;
        if (overhang > allowedOverhang + VALUE_TOL) {
          fb.push(`Vertical tail extends ${overhang.toFixed(2)} ft beyond the fuselage end (limit ${allowedOverhang.toFixed(2)} ft).`);
          failures += 1;
        }
      }
    }
  }

  const engine_length = getNumber(main, "I29");
  if (!Number.isFinite(engine_diameter) || !Number.isFinite(fuselage_end) || !Number.isFinite(inlet_x) || !Number.isFinite(compressor_x) || !Number.isFinite(engine_length)) {
    fb.push("Unable to verify engine protrusion due to missing geometry data");
    failures += 1;
  } else {
    const protrusion = inlet_x + compressor_x + engine_length - fuselage_end;
    if (protrusion > engine_diameter + VALUE_TOL) {
      fb.push(`Engine nacelles protrude ${protrusion.toFixed(2)} ft past the fuselage end (limit ${engine_diameter.toFixed(2)} ft).`);
      failures += 1;
    }
  }

  if (failures > 0) {
    fb.push(`-${Math.min(2, failures)} pts Control surface placement issues`);
  }
  return { pass: failures === 0, failures, feedback: fb };
}

function checkConstraints(main, consts, betaExpected) {
  const fb = [];
  let tableErrors = 0;
  const add = (msg) => {
    fb.push(msg);
    tableErrors += 1;
  };

  const rowCheck = (label, row, exp) => {
    const r = row - 1;
    const alt = asNumber(main?.[r]?.[19]);
    const mach = asNumber(main?.[r]?.[20]);
    const n = asNumber(main?.[r]?.[21]);
    const ab = asNumber(main?.[r]?.[22]);
    const ps = asNumber(main?.[r]?.[23]);
    const cdx = asNumber(main?.[r]?.[24]);
    const beta = asNumber(main?.[r]?.[18]);
    if (exp.altEq !== undefined) {
      if (!Number.isFinite(alt) || Math.abs(alt - exp.altEq) > TOL.alt) add(`${label}: Altitude must be ${exp.altEq} (found ${roundToTenth(alt)})`);
    } else if (exp.altMin !== undefined) {
      if (!Number.isFinite(alt) || alt < exp.altMin - TOL.alt) add(`${label}: Altitude must be >= ${exp.altMin} (found ${roundToTenth(alt)})`);
    }
    if (exp.machEq !== undefined) {
      if (!Number.isFinite(mach) || Math.abs(mach - exp.machEq) > TOL.mach) {
        if (label === "Ps2") add(`${label}: Mach = ${matlabFixed(mach, 2)}, expected ${matlabFixed(exp.machEq, 2)}`);
        else add(`${label}: Mach must be ${exp.machEq} (found ${roundToTenth(mach)})`);
      }
    } else if (exp.machMin !== undefined) {
      if (!Number.isFinite(mach) || mach < exp.machMin - TOL.mach) add(`${label}: Mach must be >= ${exp.machMin} (found ${roundToTenth(mach)})`);
    }
    if (exp.nEq !== undefined) {
      if (!Number.isFinite(n) || Math.abs(n - exp.nEq) > TOL.eq) add(`${label}: n must be ${exp.nEq.toFixed(3)} (found ${roundToTenth(n)})`);
    } else if (exp.nMin !== undefined) {
      if (!Number.isFinite(n) || n < exp.nMin - TOL.eq) add(`${label}: n must be >= ${exp.nMin.toFixed(3)} (found ${roundToTenth(n)})`);
    }
    if (!Number.isFinite(ab) || Math.abs(ab - exp.ab) > TOL.eq) add(`${label}: AB must be ${exp.ab}% (found ${roundToTenth(ab)}%)`);
    if (exp.psEq !== undefined) {
      if (!Number.isFinite(ps) || Math.abs(ps - exp.psEq) > TOL.eq) add(`${label}: Ps must be ${exp.psEq.toFixed(0)} (found ${roundToTenth(ps)})`);
    } else if (exp.psMin !== undefined) {
      if (!Number.isFinite(ps) || ps < exp.psMin - TOL.eq) add(`${label}: Ps must be >= ${exp.psMin.toFixed(0)} (found ${roundToTenth(ps)})`);
    }
    if (exp.cdxEq !== undefined) {
      if (!Number.isFinite(cdx) || Math.abs(cdx - exp.cdxEq) > TOL.eq) add(`${label}: CDx must be ${exp.cdxEq.toFixed(3)} (found ${roundToTenth(cdx)})`);
    } else if (exp.cdxAllowed !== undefined) {
      const match = Number.isFinite(cdx) && exp.cdxAllowed.some((allowed) => Math.abs(cdx - allowed) <= TOL.eq);
      if (!match) add(`${label}: CDx must be one of ${exp.cdxAllowed.map((v) => v.toFixed(3).replace(/\.?0+$/, "")).join(", ")} (found ${roundToTenth(cdx)})`);
    }
    if (!Number.isFinite(beta) || Math.abs(beta - betaExpected) > TOL.wto) add(`${label}: W/WTO must be set for 50% fuel load (${betaExpected.toFixed(3)}); found ${Number.isFinite(beta) ? beta.toFixed(3) : roundToTenth(beta)}`);
  };

  rowCheck("MaxMach", 3, { altMin: 35000, machMin: 2.0, nEq: 1, ab: 100, psEq: 0, cdxEq: 0 });
  rowCheck("CruiseMach", 4, { altMin: 35000, machMin: 1.5, nEq: 1, ab: 0, psEq: 0, cdxEq: 0 });
  rowCheck("Cmbt Turn1", 6, { altEq: 30000, machEq: 1.2, nMin: 3.0, ab: 100, psEq: 0, cdxEq: 0 });
  rowCheck("Cmbt Turn2", 7, { altEq: 10000, machEq: 0.9, nMin: 4.0, ab: 100, psEq: 0, cdxEq: 0 });
  rowCheck("Ps1", 8, { altEq: 30000, machEq: 1.15, nEq: 1, ab: 100, psMin: 400, cdxEq: 0 });
  rowCheck("Ps2", 9, { altEq: 10000, machEq: 0.9, nEq: 1, ab: 0, psMin: 400, cdxEq: 0 });

  const takeoffCdx = asNumber(main?.[11]?.[24]);
  const landingCdx = asNumber(main?.[12]?.[24]);
  const takeoffDist = getNumber(main, "X12");
  const landingDist = getNumber(main, "X13");

  // Takeoff row
  if (Math.abs(asNumber(main?.[11]?.[19])) > TOL.alt) add("Takeoff: Altitude must be 0 (found non-zero)");
  if (Math.abs(asNumber(main?.[11]?.[20]) - 1.2) > TOL.mach) add(`Takeoff: V/Vstall must be 1.2 (found ${roundToTenth(asNumber(main?.[11]?.[20]))})`);
  if (Math.abs(asNumber(main?.[11]?.[21]) - 0.03) > 5e-4) add(`Takeoff: mu must be 0.03 (found ${roundToTenth(asNumber(main?.[11]?.[21]))})`);
  if (Math.abs(asNumber(main?.[11]?.[22]) - 100) > TOL.eq) add(`Takeoff: AB must be 100% (found ${roundToTenth(asNumber(main?.[11]?.[22]))}%)`);
  if (!Number.isFinite(takeoffDist) || takeoffDist > 3000 + TOL.dist) add(`Takeoff distance exceeds threshold (3000 ft): ${takeoffDist?.toFixed?.(0) ?? roundToTenth(takeoffDist)}`);
  else if (takeoffDist <= 2500 + TOL.dist) fb.push(`Takeoff distance meets objective (<= 2500 ft): ${takeoffDist.toFixed(0)}`);
  if (Math.abs(asNumber(main?.[11]?.[18]) - 1) > TOL.wto) add(`Takeoff: W/WTO must be 1.000 within ±${TOL.wto.toFixed(3)} (found ${roundToTenth(asNumber(main?.[11]?.[18]))})`);
  if (!Number.isFinite(takeoffCdx) || ![0, 0.035].some((allowed) => Math.abs(takeoffCdx - allowed) <= TOL.eq)) add(`Takeoff: CDx must be 0 or 0.035 (found ${roundToTenth(takeoffCdx)})`);

  // Landing row
  if (Math.abs(asNumber(main?.[12]?.[19])) > TOL.alt) add("Landing: Altitude must be 0 (found non-zero)");
  if (Math.abs(asNumber(main?.[12]?.[20]) - 1.3) > TOL.mach) add(`Landing: V/Vstall must be 1.3 (found ${roundToTenth(asNumber(main?.[12]?.[20]))})`);
  if (Math.abs(asNumber(main?.[12]?.[21]) - 0.5) > TOL.eq) add(`Landing: mu must be 0.5 (found ${roundToTenth(asNumber(main?.[12]?.[21]))})`);
  if (Math.abs(asNumber(main?.[12]?.[22])) > TOL.eq) add(`Landing: AB must be 0% (found ${roundToTenth(asNumber(main?.[12]?.[22]))}%)`);
  if (!Number.isFinite(landingDist) || landingDist > 5000 + TOL.dist) add(`Landing distance exceeds threshold (5000 ft): ${landingDist?.toFixed?.(0) ?? roundToTenth(landingDist)}`);
  else if (landingDist <= 3500 + TOL.dist) fb.push(`Landing distance meets objective (<= 3500 ft): ${landingDist.toFixed(0)}`);
  if (Math.abs(asNumber(main?.[12]?.[18]) - 1) > TOL.wto) add(`Landing: W/WTO must be 1.000 within ±${TOL.wto.toFixed(3)} (found ${roundToTenth(asNumber(main?.[12]?.[18]))})`);
  if (!Number.isFinite(landingCdx) || ![0, 0.045].some((allowed) => Math.abs(landingCdx - allowed) <= TOL.eq)) add(`Landing: CDx must be 0 or 0.045 (found ${roundToTenth(landingCdx)})`);

  // Constraint curves
  let curveFailures = 0;
  const failedCurves = [];
  const curveDiagnostics = [];
  try {
    const WS_axis = (consts?.[21] ?? []).slice(10, 31).map(asNumber);
    const rows = [
      { row: 23, label: "MaxMach" },
      { row: 24, label: "Supercruise" },
      { row: 26, label: "CombatTurn1" },
      { row: 27, label: "CombatTurn2" },
      { row: 28, label: "Ps1" },
      { row: 29, label: "Ps2" },
      { row: 32, label: "Takeoff" },
    ];
    const WS_design = asNumber(main?.[12]?.[15]);
    const TW_design = asNumber(main?.[12]?.[16]);
    rows.forEach(({ row, label }) => {
      const TW_curve = (consts?.[row - 1] ?? []).slice(10, 31).map(asNumber);
      const est = pchip(WS_axis, TW_curve, WS_design);
      if (est !== null && Number.isFinite(TW_design) && TW_design < est - TOL.eq) {
        curveFailures += 1;
        failedCurves.push(label);
        curveDiagnostics.push(`Constraint curve ${label}: T/W=${roundToTenth(TW_design)} below required ${roundToTenth(est)} at W/S=${roundToTenth(WS_design)}`);
      }
    });
    const WS_limit_landing = asNumber(consts?.[32]?.[11]);
    if (Number.isFinite(WS_design) && Number.isFinite(WS_limit_landing) && WS_design > WS_limit_landing) {
      curveFailures += 1;
      failedCurves.push("Landing");
    }
  } catch (err) {
    fb.push(`Could not perform constraint curve check due to error: ${err.message}`);
    curveFailures = 0;
  }

  if (tableErrors > 0) {
    fb.push(`-${Math.min(2, tableErrors)} pts One or more constraint table entries are incorrect`);
  }

  fb.push(...curveDiagnostics);

  if (failedCurves.includes("Landing")) {
    const WS_design = asNumber(main?.[12]?.[15]);
    const WS_limit_landing = asNumber(consts?.[32]?.[11]);
    if (Number.isFinite(WS_design) && Number.isFinite(WS_limit_landing)) {
      fb.push(`Landing constraint violated: W/S = ${WS_design.toFixed(2)} exceeds limit of ${WS_limit_landing.toFixed(2)}`);
    }
  }

  if (curveFailures === 1) {
    fb.push(`-4 pts Design did not meet the following constraint: ${failedCurves[0]}. Your design is not above those limits; increase T/W or relax the offending constraint values toward their thresholds.`);
  } else if (curveFailures >= 2) {
    const suffix = curveFailures > 6 ? " Consider seeking EI; multiple constraints remain unmet." : "";
    fb.push(`-8 pts Design did not meet the following constraints: ${failedCurves.join(", ")}. Your design is not above those limits; increase T/W or relax the offending constraint values toward their thresholds.${suffix}`);
  }

  const objectiveSet = {
    maxMach: Number.isFinite(asNumber(main?.[2]?.[20])) && asNumber(main?.[2]?.[20]) >= 2.2 - TOL.mach,
    supercruise: Number.isFinite(asNumber(main?.[3]?.[20])) && asNumber(main?.[3]?.[20]) >= 1.8 - TOL.mach,
    gHigh: Number.isFinite(asNumber(main?.[5]?.[21])) && asNumber(main?.[5]?.[21]) >= 4.0 - TOL.eq,
    gLow: Number.isFinite(asNumber(main?.[6]?.[21])) && asNumber(main?.[6]?.[21]) >= 4.5 - TOL.eq,
    psHigh: Number.isFinite(asNumber(main?.[7]?.[23])) && asNumber(main?.[7]?.[23]) >= 500 - TOL.eq,
    psLow: Number.isFinite(asNumber(main?.[8]?.[23])) && asNumber(main?.[8]?.[23]) >= 500 - TOL.eq,
    takeoff: Number.isFinite(takeoffDist) && takeoffDist <= 2500 + TOL.dist,
    landing: Number.isFinite(landingDist) && landingDist <= 3500 + TOL.dist,
  };

  const rowErrorsMap = {
    maxMach: fb.some((msg) => msg.startsWith("MaxMach:")),
    supercruise: fb.some((msg) => msg.startsWith("CruiseMach:")),
    gHigh: fb.some((msg) => msg.startsWith("Cmbt Turn1:")),
    gLow: fb.some((msg) => msg.startsWith("Cmbt Turn2:")),
    psHigh: fb.some((msg) => msg.startsWith("Ps1:")),
    psLow: fb.some((msg) => msg.startsWith("Ps2:")),
  };
  const curveStatus = {
    maxMach: !failedCurves.includes("MaxMach"),
    supercruise: !failedCurves.includes("Supercruise"),
    gHigh: !failedCurves.includes("CombatTurn1"),
    gLow: !failedCurves.includes("CombatTurn2"),
    psHigh: !failedCurves.includes("Ps1"),
    psLow: !failedCurves.includes("Ps2"),
  };

  if (objectiveSet.maxMach) {
    if (curveStatus.maxMach && !rowErrorsMap.maxMach) fb.push("Constraint MaxMach set above threshold and satisfied.");
    else if (!curveStatus.maxMach) fb.push("Constraint MaxMach set at or above objective. Design fails to meet this constraint; consider lowering it toward the threshold value.");
  }
  if (objectiveSet.supercruise) {
    if (curveStatus.supercruise && !rowErrorsMap.supercruise) fb.push("Constraint CruiseMach set above threshold and satisfied.");
    else if (!curveStatus.supercruise) fb.push("Constraint CruiseMach set at or above objective. Design fails to meet this constraint; consider lowering it toward the threshold value.");
  }
  if (objectiveSet.gHigh) {
    if (curveStatus.gHigh && !rowErrorsMap.gHigh) fb.push("Constraint Cmbt Turn1 set above threshold and satisfied.");
    else if (!curveStatus.gHigh) fb.push("Constraint Cmbt Turn1 set at or above objective. Design fails to meet this constraint; consider lowering it toward the threshold value.");
  }
  if (objectiveSet.gLow) {
    if (curveStatus.gLow && !rowErrorsMap.gLow) fb.push("Constraint Cmbt Turn2 set above threshold and satisfied.");
    else if (!curveStatus.gLow) fb.push("Constraint Cmbt Turn2 set at or above objective. Design fails to meet this constraint; consider lowering it toward the threshold value.");
  }
  if (objectiveSet.psHigh) {
    if (curveStatus.psHigh && !rowErrorsMap.psHigh) fb.push("Constraint Ps1 set above threshold and satisfied.");
    else if (!curveStatus.psHigh) fb.push("Constraint Ps1 set at or above objective. Design fails to meet this constraint; consider lowering it toward the threshold value.");
  }
  if (objectiveSet.psLow) {
    if (curveStatus.psLow && !rowErrorsMap.psLow) fb.push("Constraint Ps2 set above threshold and satisfied.");
    else if (!curveStatus.psLow) fb.push("Constraint Ps2 set at or above objective. Design fails to meet this constraint; consider lowering it toward the threshold value.");
  }

  return { pass: tableErrors === 0 && curveFailures === 0, tableErrors, curveFailures, objectiveSet, feedback: fb };
}

function checkPayload(main) {
  const aim120 = getNumber(main, "AB3");
  const aim9 = getNumber(main, "AB4");
  let payloadPass = false;
  let payloadObjectivePass = false;
  const fb = [];
  if (!Number.isFinite(aim120) || aim120 < 8 - TOL.eq) {
    const count = Number.isFinite(aim120) ? aim120 : 0;
    fb.push(`-4 pts Payload must include at least 8 AIM-120Ds (found ${count.toFixed(0)})`);
  } else {
    payloadPass = true;
    if (Number.isFinite(aim9) && aim9 >= 2 - TOL.eq) {
      payloadObjectivePass = true;
      fb.push(`Payload meets objective: ${aim120.toFixed(0)} AIM-120s + ${aim9.toFixed(0)} AIM-9s`);
    }
  }
  return { payloadPass, payloadObjectivePass, feedback: fb };
}

function checkStability(main) {
  const fb = [];
  let failures = 0;
  const SM = getNumber(main, "M10");
  const clb = getNumber(main, "O10");
  const cnb = getNumber(main, "P10");
  const rat = getNumber(main, "Q10");
  if (!(SM >= -0.1 && SM <= 0.11)) {
    fb.push(`Static margin out of bounds (M10 = ${Number.isFinite(SM) ? SM.toFixed(3) : roundToTenth(SM)})`);
    failures += 1;
    if (Number.isFinite(SM) && SM < 0) fb.push("Warning: aircraft is statically unstable (SM < 0)");
  }
  if (!(clb < -0.001)) {
    fb.push(`Clb must be < -0.001 (O10 = ${clb?.toFixed?.(6) ?? "NaN"})`);
    failures += 1;
  }
  if (!(cnb > 0.002)) {
    fb.push(`Cnb must be > 0.002 (P10 = ${cnb?.toFixed?.(6) ?? "NaN"})`);
    failures += 1;
  }
  if (!(Math.abs(rat) >= 0.3 && Math.abs(rat) <= 1)) {
    fb.push(`Cnb/Clb ratio magnitude must be between 0.3 and 1.0 (Q10 = ${Number.isFinite(rat) ? rat.toFixed(3) : roundToTenth(rat)})`);
    failures += 1;
  }
  if (failures > 0) fb.push(`-${Math.min(3, failures)} pts Stability parameters outside limits`);
  return { pass: failures === 0, failures, feedback: fb };
}

function checkFuelVolume(main) {
  const fb = [];
  const fuel_available = getNumber(main, "O18");
  const fuel_required = getNumber(main, "X40");
  const volume_remaining = getNumber(main, "Q23");
  let fuelPass = true;
  let volumePass = true;
  if (!Number.isFinite(fuel_available) || !Number.isFinite(fuel_required)) {
    fb.push("Fuel check could not be evaluated because O18 or X40 is not numeric.");
    fuelPass = false;
  } else if (fuel_available + TOL.fuel < fuel_required) {
    fb.push(`Fuel available (${roundToTenth(fuel_available)}) is less than required (${roundToTenth(fuel_required)}); check reserves.`);
    fuelPass = false;
  }
  if (!Number.isFinite(volume_remaining) || volume_remaining <= 0) {
    fb.push(`-2 pts Volume remaining must be positive (Q23 = ${Number.isFinite(volume_remaining) ? volume_remaining.toFixed(2) : roundToTenth(volume_remaining)})`);
    volumePass = false;
  }
  return { fuelPass, volumePass, feedback: fb };
}

function checkCost(main) {
  const fb = [];
  const numaircraft = getNumber(main, "N31");
  const cost = getNumber(main, "Q31");
  let costPass = false;
  let costObjectivePass = false;
  if (!Number.isFinite(numaircraft)) {
    fb.push(`Number of aircraft (N31) must be 187 or 800 (found ${roundToTenth(numaircraft)}).`);
  } else if (Math.abs(numaircraft - 187) < 1e-3) {
    if (!Number.isFinite(cost)) {
      fb.push("Recurring cost missing for 187-aircraft estimate.");
    } else {
      if (cost <= 115 + TOL.eq) {
        costPass = true;
      } else {
        fb.push(`-5 pts Recurring cost exceeds threshold ($115M): $${roundToTenth(cost)}M`);
      }
      if (cost <= 100 + TOL.eq) {
        costObjectivePass = true;
        fb.push(`Recurring cost meets objective (<= $100M): $${cost.toFixed(1)}M`);
      }
    }
  } else if (Math.abs(numaircraft - 800) < 1e-3) {
    if (!Number.isFinite(cost)) {
      fb.push("Recurring cost missing for 800-aircraft estimate.");
    } else {
      if (cost <= 75 + TOL.eq) {
        costPass = true;
      } else {
        fb.push(`-5 pts Recurring cost exceeds threshold ($75M): $${roundToTenth(cost)}M`);
      }
      if (cost <= 61 + TOL.eq) {
        costObjectivePass = true;
        fb.push(`Recurring cost meets objective (<= $61M): $${cost.toFixed(1)}M`);
      }
    }
  } else {
    fb.push(`Number of aircraft (N31) must be 187 or 800 (found ${roundToTenth(numaircraft)}).`);
  }
  return { costPass, costObjectivePass, feedback: fb };
}

function checkGear(gear) {
  const fb = [];
  let failures = 0;
  let takeoffSpeedPass = true;
  const g90 = asNumber(gear?.[19]?.[9]); // J20
  if (!Number.isFinite(g90) || g90 < 80 - TOL.eq || g90 > 90.5 + TOL.eq) {
    failures += 1;
    fb.push(`Violates main gear 90/10 rule share at J20: ${fmt1(g90)}% (must be between 80.0% and 90.5%)`);
  }

  const tipbackActual = asNumber(gear?.[19]?.[11]);
  const tipbackLimit = asNumber(gear?.[20]?.[11]);
  if (!Number.isFinite(tipbackActual) || !Number.isFinite(tipbackLimit) || tipbackActual >= tipbackLimit - 1e-2) {
    failures += 1;
    fb.push(`Violates tipback angle requirement: upper ${roundToTenth(tipbackActual)}° must be less than lower ${roundToTenth(tipbackLimit)}°`);
  }

  const rolloverActual = asNumber(gear?.[19]?.[12]);
  const rolloverLimit = asNumber(gear?.[20]?.[12]);
  if (!Number.isFinite(rolloverActual) || !Number.isFinite(rolloverLimit) || rolloverActual >= rolloverLimit - 1e-2) {
    failures += 1;
    fb.push(`Violates rollover angle requirement: upper ${roundToTenth(rolloverActual)}° must be less than lower ${roundToTenth(rolloverLimit)}°`);
  }

  const rotationSpeed = asNumber(gear?.[19]?.[13]);
  const rotationRef = asNumber(gear?.[20]?.[13]);
  if (!Number.isFinite(rotationSpeed)) {
    failures += 1;
    fb.push("Takeoff rotation speed (N20) missing; N20 must be less than N21.");
    takeoffSpeedPass = false;
  } else {
    if (!Number.isFinite(rotationRef)) {
      failures += 1;
      fb.push("Takeoff speed margin failed: N21 missing; N20 must be below N21.");
      takeoffSpeedPass = false;
    } else {
      if (rotationSpeed >= 200 - TOL.eq) {
        failures += 1;
        fb.push(`Violates takeoff rotation speed: ${fmt1(rotationSpeed)} kts (must be < 200 kts)`);
        takeoffSpeedPass = false;
      }
      if (rotationSpeed >= rotationRef) {
        failures += 1;
        fb.push(`Takeoff speed margin failed: N20 must be less than N21 (N20 = ${rotationSpeed.toFixed(2)}, N21 = ${rotationRef.toFixed(2)})`);
        takeoffSpeedPass = false;
      }
      if (rotationRef > 200 + TOL.eq) {
        failures += 1;
        fb.push(`Takeoff speed too high: N21 = ${fmt1(rotationRef)} kts (must be <= 200 kts)`);
        takeoffSpeedPass = false;
      }
    }
  }

  if (failures > 0) fb.push(`-${Math.min(4, failures)} pts Landing gear geometry outside limits`);
  return { pass: failures === 0, feedback: fb, failures, takeoffSpeedPass };
}

export function gradeWorkbook(workbook) {
  const feedback = [];

  // Preflight error cells
  const invalidCells = checkInvalidMainCells(workbook.sheets?.main ?? []);
  if (invalidCells.length > 0) {
    const msg = `Invalid for analysis: Excel errors in Main sheet at ${invalidCells.join(", ")}. Correct the errors and resubmit.`;
    return { score: 0, maxScore: BASE_TOTAL, scoreLine: msg, bonusLine: "", feedbackLog: msg };
  }

  const main = workbook.sheets.main;
  const aero = workbook.sheets.aero;
  const miss = workbook.sheets.miss;
  const consts = workbook.sheets.consts;
  const gear = workbook.sheets.gear;
  const geom = workbook.sheets.geom;

  if (workbook.fileName) feedback.push(workbook.fileName);

  // Geometry blocks must be numeric
  const missingGeom = checkGeometryBlocks(main);
  if (missingGeom.missing1.length > 0) {
    const msg = `Sheet validation: Geometry inputs B18:H27 must be numeric (missing at ${missingGeom.missing1.join(", ")}).`;
    const log = [feedback[0] ?? "", msg].filter(Boolean).join("\n");
    return { score: 0, maxScore: BASE_TOTAL, scoreLine: msg, bonusLine: "", feedbackLog: log };
  }
  if (missingGeom.missing2.length > 0) {
    const msg = `Sheet validation: Geometry inputs C34:F53 must be numeric (missing at ${missingGeom.missing2.join(", ")}).`;
    const log = [feedback[0] ?? "", msg].filter(Boolean).join("\n");
    return { score: 0, maxScore: BASE_TOTAL, scoreLine: msg, bonusLine: "", feedbackLog: log };
  }

  const fuel_available = getNumber(main, "O18");
  const fuel_capacity = getNumber(main, "O15");
  const fuel_required = getNumber(main, "X40");
  const volume_remaining = getNumber(main, "Q23");
  const cost = getNumber(main, "Q31");
  const numaircraft = getNumber(main, "N31");
  const radius = getNumber(main, "Y37");
  const aim120 = getNumber(main, "AB3");
  const aim9 = getNumber(main, "AB4");
  const takeoff_dist = getNumber(main, "X12");
  const landing_dist = getNumber(main, "X13");
  const betaDefault =
    Number.isFinite(fuel_available) && Number.isFinite(fuel_capacity) && fuel_capacity !== 0
      ? 1 - fuel_available / (2 * fuel_capacity)
      : BETA_DEFAULT;

  const aeroResult = checkAero(aero);
  feedback.push(...aeroResult.feedback);

  const mission = checkMissionProfile(main, radius, betaDefault);
  feedback.push(...mission.feedback);
  if (!mission.missionPass) {
    feedback.push(`-${Math.min(2, mission.missionErrors)} pts Mission profile inputs incorrect`);
  }

  const efficiency = checkEfficiency(main);

  const thrust = checkThrust(miss);
  feedback.push(...thrust.feedback);

  const control = checkControlAttachment(main, geom);
  feedback.push(...control.feedback);

  const stealthResult = runStealthChecks(workbook);
  const stealthPass = stealthResult.failures === 0;
  feedback.push(...stealthResult.feedback);

  const constraints = checkConstraints(main, consts, mission.betaExpected);
  feedback.push(...constraints.feedback);

  const payload = checkPayload(main);
  feedback.push(...payload.feedback);

  const stability = checkStability(main);
  feedback.push(...stability.feedback);

  const fuelVolume = checkFuelVolume(main);
  feedback.push(...fuelVolume.feedback);

  const costResult = checkCost(main);
  feedback.push(...costResult.feedback);

  const gearResult = checkGear(gear);
  feedback.push(...gearResult.feedback);

  let pt = BASE_TOTAL;
  pt -= aeroResult.deduction;
  if (!thrust.pass) pt -= 3;
  if (!mission.missionPass) pt -= Math.min(2, mission.missionErrors);
  if (constraints.tableErrors > 0) pt -= Math.min(2, constraints.tableErrors);
  if (constraints.curveFailures === 1) pt -= 4;
  else if (constraints.curveFailures >= 2) pt -= 8;
  if (!payload.payloadPass) pt -= 4;
  if (!control.pass) pt -= Math.min(2, control.failures);
  if (!stability.pass) pt -= Math.min(3, stability.failures);
  if (!fuelVolume.fuelPass) pt -= 2;
  if (!fuelVolume.volumePass) pt -= 2;
  if (!costResult.costPass) pt -= 5;
  if (!gearResult.pass) pt -= Math.min(4, gearResult.failures);
  pt -= stealthResult.deduction;

  pt = Math.max(0, pt);

  const radiusBonus = roundToTenth(linearBonus(radius, 375, 410));
  const payloadBonus = roundToTenth(Number.isFinite(aim120) && Number.isFinite(aim9) && aim120 >= 8 - TOL.eq && aim9 >= 2 - TOL.eq ? 1 : 0);
  const takeoffBonus = roundToTenth(linearBonusInv(takeoff_dist, 3000, 2500));
  const landingBonus = roundToTenth(linearBonusInv(landing_dist, 5000, 3500));
  const maxMachBonus = roundToTenth(linearBonus(asNumber(main?.[2]?.[20]), 2.0, 2.2));
  const superBonus = roundToTenth(linearBonus(asNumber(main?.[3]?.[20]), 1.5, 1.8));
  const psHighBonus = roundToTenth(linearBonus(asNumber(main?.[7]?.[23]), 400, 500));
  const psLowBonus = roundToTenth(linearBonus(asNumber(main?.[8]?.[23]), 400, 500));
  const gHighBonus = roundToTenth(linearBonus(asNumber(main?.[5]?.[21]), 3.0, 4.0));
  const gLowBonus = roundToTenth(linearBonus(asNumber(main?.[6]?.[21]), 4.0, 4.5));
  let costBonus = 0;
  if (Math.abs(numaircraft - 187) < 1e-3) costBonus = roundToTenth(linearBonusInv(cost, 115, 100));
  else if (Math.abs(numaircraft - 800) < 1e-3) costBonus = roundToTenth(linearBonusInv(cost, 75, 61));

  const bonusEligible = constraints.tableErrors === 0 && constraints.curveFailures === 0;
  let objectiveScore =
    radiusBonus +
    payloadBonus +
    takeoffBonus +
    landingBonus +
    maxMachBonus +
    superBonus +
    psHighBonus +
    psLowBonus +
    gHighBonus +
    gLowBonus +
    costBonus;

  if (!bonusEligible) {
    objectiveScore = 0;
    feedback.push("Bonus points unavailable because one or more constraints miss threshold values or the design is below a constraint curve.");
  } else {
    if (radiusBonus > 0) feedback.push(`Mission radius bonus [+${radiusBonus.toFixed(1)} bonus]: ${radius.toFixed(1)} nm`);
    if (payloadBonus > 0) feedback.push(`Payload bonus [+${payloadBonus.toFixed(1)} bonus]: ${aim120.toFixed(0)} AIM-120s + ${aim9.toFixed(0)} AIM-9s`);
    if (takeoffBonus > 0) feedback.push(`Takeoff distance bonus [+${takeoffBonus.toFixed(1)} bonus]: ${takeoff_dist.toFixed(0)} ft`);
    if (landingBonus > 0) feedback.push(`Landing distance bonus [+${landingBonus.toFixed(1)} bonus]: ${landing_dist.toFixed(0)} ft`);
    if (maxMachBonus > 0) feedback.push(`Max Mach bonus [+${maxMachBonus.toFixed(1)} bonus]: Mach ${matlabFixed(asNumber(main?.[2]?.[20]), 2)}`);
    if (superBonus > 0) feedback.push(`Supercruise Mach bonus [+${superBonus.toFixed(1)} bonus]: Mach ${matlabFixed(asNumber(main?.[3]?.[20]), 2)}`);
    if (psHighBonus > 0) feedback.push(`Ps @30k ft bonus [+${psHighBonus.toFixed(1)} bonus]: ${asNumber(main?.[7]?.[23]).toFixed(0)} ft/s`);
    if (psLowBonus > 0) feedback.push(`Ps @10k ft bonus [+${psLowBonus.toFixed(1)} bonus]: ${asNumber(main?.[8]?.[23]).toFixed(0)} ft/s`);
    if (gHighBonus > 0) feedback.push(`Combat turn (30k ft) bonus [+${gHighBonus.toFixed(1)} bonus]: ${asNumber(main?.[5]?.[21]).toFixed(2)} g`);
    if (gLowBonus > 0) feedback.push(`Combat turn (10k ft) bonus [+${gLowBonus.toFixed(1)} bonus]: ${asNumber(main?.[6]?.[21]).toFixed(2)} g`);
    if (costBonus > 0) feedback.push(`Recurring cost bonus [+${costBonus.toFixed(1)} bonus]: ${numaircraft.toFixed(0)} aircraft, $${cost.toFixed(1)}M`);
  }

  const thresholdScore = roundToTenth(pt);
  objectiveScore = roundToTenth(objectiveScore);
  pt = roundToTenth(thresholdScore + objectiveScore);

  const nonViableReasons = [];
  if (!fuelVolume.fuelPass) nonViableReasons.push("fuel check failed");
  if (!fuelVolume.volumePass) nonViableReasons.push("insufficient volume remaining");
  if (!gearResult.takeoffSpeedPass) nonViableReasons.push("takeoff speed check failed");
  if (nonViableReasons.length > 0) {
    pt = roundToTenth(Math.min(pt, NON_VIABLE_CAP));
    feedback.push(`Final score capped at ${NON_VIABLE_CAP.toFixed(1)} out of ${BASE_TOTAL} because the aircraft is non-viable (${nonViableReasons.join(", ")}).`);
  }

  const scoreSummary = [
    `Jet11 base score: ${formatScore(thresholdScore)} out of ${BASE_TOTAL}`,
    `Bonus points: +${objectiveScore.toFixed(1)} (final score ${pt.toFixed(1)})`,
  ];
  feedback.push(...scoreSummary);

  const feedbackLog = feedback.join("\n");

  return {
    score: pt,
    maxScore: BASE_TOTAL,
    scoreLine: scoreSummary[0],
    bonusLine: scoreSummary[1],
    feedbackLog,
  };
}
