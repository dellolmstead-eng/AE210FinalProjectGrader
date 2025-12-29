import { getCell, asNumber } from "./parseUtils.js";
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
  wto: 1e-2,
  alt: 1,
  mach: 1e-2,
  time: 1e-2,
  dist: 1e-3,
};

const BETA_DEFAULT = 0.87620980519917;
const BASE_TOTAL = 85;
const OBJECTIVE_TOTAL = 15;

const roundToTenth = (value) => (Number.isFinite(value) ? Math.round(value * 10) / 10 : 0);
const ternary = (cond, a, b) => (cond ? a : b);

const getNumber = (sheet, ref) => asNumber(getCell(sheet, ref));

const checkInvalidMainCells = (mainSheet) => {
  const invalidCells = [];
  mainSheet?.forEach((row, rIdx) => {
    if (!row) return;
    row.forEach((value, cIdx) => {
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
  const missing = [];
  const pushMissing = (row, col) => missing.push(cellRef(row - 1, col - 1));

  // Block B18:H27 with skips
  const skips1 = new Set(["B24", "C24", "D27", "E27", "F27", "G27", "H26"]);
  for (let r = 18; r <= 27; r += 1) {
    for (let c = 2; c <= 8; c += 1) {
      const ref = `${String.fromCharCode(64 + c)}${r}`;
      if (skips1.has(ref)) continue;
      if (!Number.isFinite(getNumber(main, ref))) pushMissing(r, c);
    }
  }

  // Block C34:F53
  for (let r = 34; r <= 53; r += 1) {
    for (let c = 3; c <= 6; c += 1) {
      const ref = `${String.fromCharCode(64 + c)}${r}`;
      if (!Number.isFinite(getNumber(main, ref))) pushMissing(r, c);
    }
  }

  return missing;
}

function checkMissionProfile(main, radius, betaExpected) {
  const feedback = [];
  const colIdx = [...Array(14).keys()].map((i) => i + 1); // legs 1-14
  const MissionArray = main.slice(32, 44).map((row) => row ?? []); // rows 33-44
  const val = (rIdx, cIdx) => asNumber(MissionArray[rIdx]?.[cIdx]);

  const alt = colIdx.map((i) => val(0, i));
  const mach = colIdx.map((i) => val(2, i));
  const ab = colIdx.map((i) => val(3, i));
  const dist = colIdx.map((i) => val(5, i));
  const time = colIdx.map((i) => val(6, i));

  const altExpected = [0, 2000, 35000, 35000, 35000, 35000, 35000, 30000, 35000, 35000, 35000, 35000, 10000, 0];
  const machExpected = [0.268473504, 0.88, 0.88, 0.88, 0.88, 1.5, 0.8, 0.8, 1.5, 0.8, 0.88, 0.88, 0.4, 0.0];
  const abExpected = [100, 0, 0, 0, 0, 0, 0, 100, 0, 0, 0, 0, 0, 0];
  const supercruiseCols = new Set([6, 9]);
  const distExpected = { 6: 400, 9: 400 };
  const combatCol = 8;
  const loiterCol = 13;
  const timeExpected = { 8: 2, 13: 20 };

  let missionErrors = 0;
  colIdx.forEach((leg) => {
    const i = leg - 1;
    if (Math.abs(alt[i] - altExpected[i]) > TOL.alt) {
      feedback.push(`Leg ${leg} Altitude must be ${altExpected[i].toFixed(0)} (found ${roundToTenth(alt[i])})`);
      missionErrors += 1;
    }
    if (leg !== 1 && Math.abs(mach[i] - machExpected[i]) > TOL.mach) {
      feedback.push(`Leg ${leg} Mach must be ${machExpected[i].toFixed(2)} (found ${roundToTenth(mach[i])})`);
      missionErrors += 1;
    }
    if (leg !== 14 && Math.abs(ab[i] - abExpected[i]) > TOL.eq) {
      feedback.push(`Leg ${leg} AB must be ${abExpected[i].toFixed(0)} (found ${roundToTenth(ab[i])})`);
      missionErrors += 1;
    }
    if (supercruiseCols.has(leg)) {
      const expected = distExpected[leg];
      if (Math.abs(dist[i] - expected) > TOL.dist) {
        feedback.push(`Leg ${leg} Supercruise distance must be ${expected} (found ${roundToTenth(dist[i])})`);
        missionErrors += 1;
      }
    }
    if (leg === combatCol) {
      const expected = timeExpected[leg];
      if (time[i] < expected - TOL.time) {
        feedback.push(`Leg ${leg} Time must be >= ${expected.toFixed(2)} min (found ${roundToTenth(time[i])})`);
        missionErrors += 1;
      }
    }
    if (leg === loiterCol) {
      const expected = timeExpected[leg];
      if (Math.abs(time[i] - expected) > TOL.time) {
        feedback.push(`Leg ${leg} Time must be ${expected.toFixed(2)} min (found ${roundToTenth(time[i])})`);
        missionErrors += 1;
      }
    }
  });

  let rangePass = false;
  let rangeObjectivePass = false;
  if (Number.isFinite(radius)) {
    if (radius >= 800 - TOL.dist) {
      rangePass = true;
      rangeObjectivePass = true;
    } else if (radius >= 500 - TOL.dist) {
      rangePass = true;
    } else {
      feedback.push(`Range below threshold: mission radius = ${roundToTenth(radius)} nm (needs >= 500 nm)`);
      missionErrors += 1;
    }
  } else {
    feedback.push("Mission radius missing; unable to verify range requirement.");
    missionErrors += 1;
  }

  const missionPass = missionErrors === 0;
  return { feedback, missionPass, rangePass, rangeObjectivePass, betaExpected };
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
    const available = asNumber(miss?.[47]?.[c]);
    const drag = asNumber(miss?.[48]?.[c]);
    if (!Number.isFinite(available) || !Number.isFinite(drag)) continue;
    if (drag >= available - TOL.eq) thrustShort.push(c);
  }
  if (thrustShort.length > 0) {
    feedback.push(`Thrust shortfall: Tavailable <= Drag for ${thrustShort.length} mission segment(s).`);
  }
  return { pass: thrustShort.length === 0, feedback };
}

function checkControlAttachment(main, geom) {
  const fb = [];
  let failures = 0;
  const VALUE_TOL = 1e-3;
  const AR_TOL = 0.1;
  const VT_WING_FRACTION = 0.8;

  const fuselage_end = getNumber(main, "B32");
  const PCS_x = getNumber(main, "C23");
  const PCS_root = getNumber(geom, "C8");
  if (!Number.isFinite(fuselage_end) || !Number.isFinite(PCS_x) || !Number.isFinite(PCS_root)) {
    fb.push("Unable to verify PCS placement due to missing geometry data");
    failures += 1;
  } else if (PCS_x > fuselage_end - 0.25 * PCS_root) {
    fb.push("PCS X-location too far aft. Must overlap at least 25% of root chord.");
    failures += 1;
  }

  const VT_x = getNumber(main, "H23");
  const VT_root = getNumber(geom, "C10");
  if (!Number.isFinite(fuselage_end) || !Number.isFinite(VT_x) || !Number.isFinite(VT_root)) {
    fb.push("Unable to verify vertical tail placement due to missing geometry data");
    failures += 1;
  } else if (VT_x > fuselage_end - 0.25 * VT_root) {
    fb.push("VT X-location too far aft. Must overlap at least 25% of root chord.");
    failures += 1;
  }

  const PCS_z = getNumber(main, "C25");
  const fuse_z_center = getNumber(main, "D52");
  const fuse_z_height = getNumber(main, "F52");
  if (!Number.isFinite(PCS_z) || !Number.isFinite(fuse_z_center) || !Number.isFinite(fuse_z_height)) {
    fb.push("Unable to verify PCS vertical placement due to missing geometry data");
    failures += 1;
  } else if (PCS_z < fuse_z_center - fuse_z_height / 2 || PCS_z > fuse_z_center + fuse_z_height / 2) {
    fb.push("PCS Z-location outside fuselage vertical bounds.");
    failures += 1;
  }

  const VT_y = getNumber(main, "H24");
  const fuse_width = getNumber(main, "E52");
  let vtMountedOffFuselage = false;
  if (!Number.isFinite(VT_y) || !Number.isFinite(fuse_width)) {
    fb.push("Unable to verify vertical tail lateral placement due to missing geometry data");
    failures += 1;
  } else if (Math.abs(VT_y) > fuse_width / 2 + VALUE_TOL) {
    vtMountedOffFuselage = true;
    fb.push("Vertical tail mounted off the fuselage; ensure structural support at the wing.");
  }

  if (getNumber(main, "D18") > 1) {
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

  const component_positions = main?.[22]?.slice(1, 8).map(asNumber) || [];
  if (component_positions.some((v) => Number.isFinite(v) && v >= fuselage_end)) {
    fb.push(`One or more components X-location extend beyond the fuselage end (B32 = ${roundToTenth(fuselage_end)})`);
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
    if (!Number.isFinite(alt) || Math.abs(alt - exp.alt) > TOL.alt) add(`${label}: Altitude must be ${exp.alt} (found ${roundToTenth(alt)})`);
    if (!Number.isFinite(mach) || Math.abs(mach - exp.mach) > TOL.mach) add(`${label}: Mach must be ${exp.mach} (found ${roundToTenth(mach)})`);
    if (!Number.isFinite(n) || Math.abs(n - exp.n) > TOL.eq) add(`${label}: n must be ${exp.n.toFixed(3)} (found ${roundToTenth(n)})`);
    if (!Number.isFinite(ab) || Math.abs(ab - exp.ab) > TOL.eq) add(`${label}: AB must be ${exp.ab}% (found ${roundToTenth(ab)}%)`);
    if (!Number.isFinite(ps) || Math.abs(ps - exp.ps) > TOL.eq) add(`${label}: Ps must be ${exp.ps.toFixed(0)} (found ${roundToTenth(ps)})`);
    if (exp.cdx !== undefined) {
      if (!Number.isFinite(cdx) || Math.abs(cdx - exp.cdx) > TOL.eq) add(`${label}: CDx must be ${exp.cdx.toFixed(3)} (found ${roundToTenth(cdx)})`);
    }
    if (!Number.isFinite(beta) || Math.abs(beta - betaExpected) > TOL.wto) add(`${label}: W/WTO must be set for 50% fuel load (${betaExpected.toFixed(3)}); found ${roundToTenth(beta)}`);
  };

  rowCheck("MaxMach", 3, { alt: 35000, mach: 2.0, n: 1, ab: 100, ps: 0, cdx: 0 });
  rowCheck("CruiseMach", 4, { alt: 35000, mach: 1.5, n: 1, ab: 0, ps: 0, cdx: 0 });

  rowCheck("Supercruise", 5, { alt: 50000, mach: 1.5, n: 1, ab: 100, ps: 0, cdx: 0 });
  rowCheck("Cmbt Turn1", 6, { alt: 30000, mach: 1.2, n: 3.0, ab: 100, ps: 0, cdx: 0 });
  rowCheck("Cmbt Turn2", 7, { alt: 10000, mach: 0.9, n: 4.0, ab: 100, ps: 0, cdx: 0 });
  rowCheck("Ps1", 8, { alt: 30000, mach: 1.15, n: 1, ab: 100, ps: 400, cdx: 0 });
  rowCheck("Ps2", 9, { alt: 10000, mach: 0.9, n: 1, ab: 0, ps: 400, cdx: 0 });

  const takeoffCdx = asNumber(main?.[11]?.[24]);
  const landingCdx = asNumber(main?.[12]?.[24]);
  const takeoffDist = getNumber(main, "X12");
  const landingDist = getNumber(main, "X13");

  // Takeoff row
  if (Math.abs(asNumber(main?.[11]?.[19])) > TOL.alt) add("Takeoff: Altitude must be 0 (found non-zero)");
  if (Math.abs(asNumber(main?.[11]?.[20]) - 1.2) > TOL.mach) add(`Takeoff: V/Vstall must be 1.2 (found ${roundToTenth(asNumber(main?.[11]?.[20]))})`);
  if (Math.abs(asNumber(main?.[11]?.[21]) - 0.03) > 5e-4) add(`Takeoff: mu must be 0.03 (found ${roundToTenth(asNumber(main?.[11]?.[21]))})`);
  if (Math.abs(asNumber(main?.[11]?.[22]) - 100) > TOL.eq) add(`Takeoff: AB must be 100% (found ${roundToTenth(asNumber(main?.[11]?.[22]))}%)`);
  if (!Number.isFinite(takeoffDist) || Math.abs(takeoffDist - 3000) > TOL.dist) add(`Takeoff distance must be 3000 ft (found ${roundToTenth(takeoffDist)})`);
  if (Math.abs(asNumber(main?.[11]?.[18]) - 1) > TOL.wto) add(`Takeoff: W/WTO must be 1.000 within ±${TOL.wto.toFixed(3)} (found ${roundToTenth(asNumber(main?.[11]?.[18]))})`);
  if (!Number.isFinite(takeoffCdx) || Math.abs(takeoffCdx - 0.035) > TOL.eq) add(`Takeoff: CDx must be 0.035 (found ${roundToTenth(takeoffCdx)})`);

  // Landing row
  if (Math.abs(asNumber(main?.[12]?.[19])) > TOL.alt) add("Landing: Altitude must be 0 (found non-zero)");
  if (Math.abs(asNumber(main?.[12]?.[20]) - 1.3) > TOL.mach) add(`Landing: V/Vstall must be 1.3 (found ${roundToTenth(asNumber(main?.[12]?.[20]))})`);
  if (Math.abs(asNumber(main?.[12]?.[21]) - 0.5) > TOL.eq) add(`Landing: mu must be 0.5 (found ${roundToTenth(asNumber(main?.[12]?.[21]))})`);
  if (Math.abs(asNumber(main?.[12]?.[22])) > TOL.eq) add(`Landing: AB must be 0% (found ${roundToTenth(asNumber(main?.[12]?.[22]))}%)`);
  if (!Number.isFinite(landingDist) || Math.abs(landingDist - 5000) > TOL.dist) add(`Landing distance must be 5000 ft (found ${roundToTenth(landingDist)})`);
  if (Math.abs(asNumber(main?.[12]?.[18]) - 1) > TOL.wto) add(`Landing: W/WTO must be 1.000 within ±${TOL.wto.toFixed(3)} (found ${roundToTenth(asNumber(main?.[12]?.[18]))})`);
  if (!Number.isFinite(landingCdx) || Math.abs(landingCdx - 0.045) > TOL.eq) add(`Landing: CDx must be 0.045 (found ${roundToTenth(landingCdx)})`);

  // Constraint curves
  let curveFailures = 0;
  const failedCurves = [];
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
    const interp = (xList, yList, x) => {
      const pairs = xList
        .map((v, idx) => ({ x: asNumber(v), y: asNumber(yList[idx]) }))
        .filter((p) => Number.isFinite(p.x) && Number.isFinite(p.y))
        .sort((a, b) => a.x - b.x);
      if (pairs.length === 0) return null;
      if (x <= pairs[0].x) return pairs[0].y;
      if (x >= pairs[pairs.length - 1].x) return pairs[pairs.length - 1].y;
      for (let i = 0; i < pairs.length - 1; i += 1) {
        const p0 = pairs[i];
        const p1 = pairs[i + 1];
        if (x >= p0.x && x <= p1.x) {
          const slope = (p1.y - p0.y) / (p1.x - p0.x);
          return p0.y + slope * (x - p0.x);
        }
      }
      return null;
    };

    rows.forEach(({ row, label }) => {
      const TW_curve = (consts?.[row - 1] ?? []).slice(10, 31).map(asNumber);
      const est = interp(WS_axis, TW_curve, WS_design);
      if (est !== null && Number.isFinite(TW_design) && TW_design < est - TOL.eq) {
        curveFailures += 1;
        failedCurves.push(label);
        fb.push(`Constraint curve ${label}: T/W=${roundToTenth(TW_design)} below required ${roundToTenth(est)} at W/S=${roundToTenth(WS_design)}`);
      }
    });
    const WS_limit_landing = asNumber(consts?.[32]?.[11]);
    if (Number.isFinite(WS_design) && Number.isFinite(WS_limit_landing) && WS_design > WS_limit_landing) {
      curveFailures += 1;
      failedCurves.push("Landing");
      fb.push(`Landing constraint violated: W/S = ${roundToTenth(WS_design)} exceeds limit of ${roundToTenth(WS_limit_landing)}`);
    }
  } catch (err) {
    fb.push(`Could not perform constraint curve check due to error: ${err.message}`);
    curveFailures = 0;
  }

  if (curveFailures === 1) {
    fb.push(`Design did not meet the following constraint curve: ${failedCurves[0]}.`);
  } else if (curveFailures >= 2) {
    fb.push(`Design did not meet the following constraint curves: ${failedCurves.join(", ")}.`);
  }

  return { pass: tableErrors === 0 && curveFailures === 0, feedback: fb };
}

function checkPayload(main) {
  const aim120 = getNumber(main, "AB3");
  const aim9 = getNumber(main, "AB4");
  let payloadPass = false;
  let payloadObjectivePass = false;
  const fb = [];
  if (!Number.isFinite(aim120) || aim120 < 8 - TOL.eq) {
    const count = Number.isFinite(aim120) ? aim120 : 0;
    fb.push(`Payload missing: need at least 8 AIM-120Ds (found ${roundToTenth(count)})`);
  } else {
    payloadPass = true;
    if (Number.isFinite(aim9) && aim9 >= 2 - TOL.eq) payloadObjectivePass = true;
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
    fb.push(`Static margin out of bounds (M10 = ${roundToTenth(SM)})`);
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
  if (!(rat >= -1 && rat <= -0.3)) {
    fb.push(`Cnb/Clb ratio must be between -1 and -0.3 (Q10 = ${roundToTenth(rat)})`);
    failures += 1;
  }
  if (failures > 0) fb.push(`Stability criteria failed in ${failures} area(s).`);
  return { pass: failures === 0, feedback: fb };
}

function checkFuelVolume(main) {
  const fb = [];
  const fuel_available = getNumber(main, "O18");
  const fuel_required = getNumber(main, "X40");
  const volume_remaining = getNumber(main, "Q23");
  let fuelPass = true;
  let volumePass = true;
  if (!Number.isFinite(fuel_available) || !Number.isFinite(fuel_required) || fuel_available + TOL.eq < fuel_required) {
    fb.push(`Fuel available (${roundToTenth(fuel_available)}) is less than required (${roundToTenth(fuel_required)}); check reserves.`);
    fuelPass = false;
  }
  if (!Number.isFinite(volume_remaining) || volume_remaining <= 0) {
    fb.push(`Volume remaining must be positive (Q23 = ${roundToTenth(volume_remaining)}).`);
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
  if (!Number.isFinite(numaircraft) || Math.abs(numaircraft - 187) > 1e-3) {
    fb.push(`Number of aircraft (N31) must be 187 to evaluate cost thresholds (found ${roundToTenth(numaircraft)}).`);
  } else if (!Number.isFinite(cost)) {
    fb.push("Recurring cost missing for 187-aircraft estimate.");
  } else {
    if (cost < 120 + TOL.eq) {
      costPass = true;
    } else {
      fb.push(`Cost above threshold: $${roundToTenth(cost)}M for 187 aircraft (needs <$120M).`);
    }
    if (cost < 110 + TOL.eq) costObjectivePass = true;
  }
  return { costPass, costObjectivePass, feedback: fb };
}

function checkGear(gear) {
  const fb = [];
  let failures = 0;
  const g90 = asNumber(gear?.[19]?.[9]); // J20
  if (!Number.isFinite(g90) || g90 < 80 - TOL.eq || g90 > 95 + TOL.eq) {
    failures += 1;
    fb.push(`Violates nose gear 90/10 rule: ${roundToTenth(g90)}% (must be between 80% and 95%)`);
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
    fb.push("Takeoff rotation speed (N20) missing; must be <200 kts and below N21.");
  } else {
    if (rotationSpeed >= 200 - TOL.eq) {
      failures += 1;
      fb.push(`Violates takeoff rotation speed: N20 = ${roundToTenth(rotationSpeed)} kts (must be < 200 kts)`);
    }
    if (!Number.isFinite(rotationRef)) {
      failures += 1;
      fb.push("Takeoff speed margin failed: N21 missing; N20 must be below N21.");
    } else {
      if (rotationSpeed >= rotationRef - TOL.eq) {
        failures += 1;
        fb.push(`Takeoff speed margin failed: N20 must be less than N21 (N20 = ${roundToTenth(rotationSpeed)}, N21 = ${roundToTenth(rotationRef)})`);
      }
      if (rotationRef > 200 + TOL.eq) {
        failures += 1;
        fb.push(`Takeoff speed too high: N21 = ${roundToTenth(rotationRef)} kts (must be <= 200 kts).`);
      }
    }
  }

  if (failures > 0) fb.push(`Landing gear geometry outside limits in ${failures} area(s).`);
  return { pass: failures === 0, feedback: fb, failures };
}

export function gradeWorkbook(workbook) {
  const feedback = [];

  // Preflight error cells
  const invalidCells = checkInvalidMainCells(workbook.sheets?.main ?? []);
  if (invalidCells.length > 0) {
    const msg = `Invalid for analysis: Excel errors in Main sheet at ${invalidCells.join(", ")}. Correct the errors and resubmit.`;
    return { score: 0, maxScore: 100, scoreLine: msg, bonusLine: "", feedbackLog: msg };
  }

  const main = workbook.sheets.main;
  const miss = workbook.sheets.miss;
  const consts = workbook.sheets.consts;
  const gear = workbook.sheets.gear;
  const geom = workbook.sheets.geom;

  if (workbook.fileName) feedback.push(workbook.fileName);

  // Geometry blocks must be numeric
  const missingGeom = checkGeometryBlocks(main);
  if (missingGeom.length > 0) {
    const msg = `Sheet validation: Geometry inputs must be numeric (missing at ${missingGeom.join(", ")}).`;
    const log = [feedback[0] ?? "", msg].filter(Boolean).join("\n");
    return { score: 0, maxScore: 100, scoreLine: msg, bonusLine: "", feedbackLog: log };
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

  const mission = checkMissionProfile(main, radius, betaDefault);
  feedback.push(...mission.feedback);

  const efficiency = checkEfficiency(main);
  feedback.push(...efficiency.feedback);

  const thrust = checkThrust(miss);
  feedback.push(...thrust.feedback);

  const control = checkControlAttachment(main, geom);
  feedback.push(...control.feedback);

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

  const stealthResult = runStealthChecks(workbook);
  const stealthPass = stealthResult.feedback.length === 0;
  feedback.push(...stealthResult.feedback);

  // Scoring buckets
  let pt = BASE_TOTAL;
  const constraintsBucketPass =
    constraints.pass && payload.payloadPass && efficiency.pass && thrust.pass && missingGeom.length === 0;
  const geometryBucketPass = control.pass && stability.pass;
  const stealthBucketPass = stealthPass;

  if (!constraintsBucketPass) pt -= 5;
  if (!mission.rangePass) pt -= 5;
  if (!geometryBucketPass) pt -= 5;
  if (!gearResult.pass) pt -= 5;
  if (!fuelVolume.fuelPass) pt -= 5;
  if (!fuelVolume.volumePass) pt -= 5;
  if (!stealthBucketPass) pt -= 5;

  pt = Math.max(0, pt);

  let objectiveScore = 0;
  if (mission.rangeObjectivePass) objectiveScore += 5;
  if (costResult.costObjectivePass) objectiveScore += 5;
  if (payload.payloadObjectivePass) objectiveScore += 5;

  const thresholdScore = roundToTenth(pt);
  objectiveScore = roundToTenth(objectiveScore);
  pt = roundToTenth(thresholdScore + objectiveScore);

  const missing = [];
  if (!constraintsBucketPass) missing.push("constraints/payload/efficiency/Tavail/sheet validity");
  if (!mission.rangePass) missing.push("range");
  if (!geometryBucketPass) missing.push("geometry (controls/stability)");
  if (!gearResult.pass) missing.push("landing gear");
  if (!fuelVolume.fuelPass) missing.push("fuel");
  if (!fuelVolume.volumePass) missing.push("volume remaining");
  if (!stealthBucketPass) missing.push("stealth shaping");
  if (!mission.missionPass) missing.push("mission table (no deduction)");

  const bucketSummary = [
    "Bucket summary:",
    `  Constraints: ${ternary(constraintsBucketPass, "PASS", "FAIL (-5)")}`,
    `  Range: ${ternary(mission.rangePass, "PASS", "FAIL (-5)")}`,
    `  Geometry: ${ternary(geometryBucketPass, "PASS", "FAIL (-5)")}`,
    `  Gear: ${ternary(gearResult.pass, "PASS", "FAIL (-5)")}`,
    `  Fuel: ${ternary(fuelVolume.fuelPass, "PASS", "FAIL (-5)")}`,
    `  Volume: ${ternary(fuelVolume.volumePass, "PASS", "FAIL (-5)")}`,
    `  Stealth: ${ternary(stealthBucketPass, "PASS", "FAIL (-5)")}`,
    `Objectives: Range ${ternary(mission.rangeObjectivePass, "PASS", "FAIL")}, Cost ${ternary(
      costResult.costObjectivePass,
      "PASS",
      "FAIL"
    )}, Payload ${ternary(payload.payloadObjectivePass, "PASS", "FAIL")} => +${objectiveScore.toFixed(1)} / ${OBJECTIVE_TOTAL}`,
  ].join("\n");

  const scoreSummary = [`Threshold score after deductions: ${thresholdScore.toFixed(1)} / ${BASE_TOTAL}`, `Final score: ${pt.toFixed(1)} / 100`];

  if (missing.length > 0) {
    feedback.push(`Checks not met: ${missing.join(", ")}`);
  }
  feedback.push(bucketSummary);
  feedback.push(...scoreSummary);

  const feedbackLog = feedback.join("\n");

  return {
    score: pt,
    maxScore: 100,
    scoreLine: scoreSummary[1],
    bonusLine: "",
    feedbackLog,
  };
}
