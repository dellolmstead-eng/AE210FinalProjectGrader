import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";
import { format } from "../format.js";
import { getPlanformPoint, getMaxX } from "../geomUtils.js";

const DEG_TO_RAD = Math.PI / 180;
const VALUE_TOL = 1e-3;
const AR_TOL = 0.1;
const VT_WING_FRACTION = 0.8;

export function runAttachmentChecks(workbook) {
  const feedback = [];
  let controlFailures = 0;

  const main = workbook.sheets.main;
  const geom = workbook.sheets.geom;

  const fuselageLength = asNumber(getCell(main, "B32"));
  const pcsX = asNumber(getCell(main, "C23"));
  const pcsRootChord = asNumber(getCell(geom, "C8"));
  if (
    !Number.isFinite(fuselageLength) ||
    !Number.isFinite(pcsX) ||
    !Number.isFinite(pcsRootChord)
  ) {
    feedback.push("Unable to verify PCS placement due to missing geometry data");
    controlFailures += 1;
  } else if (pcsX > fuselageLength - 0.25 * pcsRootChord) {
    feedback.push(STRINGS.attachment.pcsX);
    controlFailures += 1;
  }

  const vtX = asNumber(getCell(main, "H23"));
  const vtRootChord = asNumber(getCell(geom, "C10"));
  if (
    !Number.isFinite(fuselageLength) ||
    !Number.isFinite(vtX) ||
    !Number.isFinite(vtRootChord)
  ) {
    feedback.push("Unable to verify vertical tail placement due to missing geometry data");
    controlFailures += 1;
  } else if (vtX > fuselageLength - 0.25 * vtRootChord) {
    feedback.push(STRINGS.attachment.vtX);
    controlFailures += 1;
  }

  const pcsZ = asNumber(getCell(main, "C25"));
  const fuseZCenter = asNumber(getCell(main, "D52"));
  const fuseZHeight = asNumber(getCell(main, "F52"));
  if (
    Number.isFinite(pcsZ) &&
    Number.isFinite(fuseZCenter) &&
    Number.isFinite(fuseZHeight)
  ) {
    if (pcsZ < fuseZCenter - fuseZHeight / 2 || pcsZ > fuseZCenter + fuseZHeight / 2) {
      feedback.push(STRINGS.attachment.pcsZ);
      controlFailures += 1;
    }
  } else {
    feedback.push("Unable to verify PCS vertical placement due to missing geometry data");
    controlFailures += 1;
  }

  const vtY = asNumber(getCell(main, "H24"));
  const fuseWidth = asNumber(getCell(main, "E52"));
  let vtMountedOffFuselage = false;
  if (Number.isFinite(vtY) && Number.isFinite(fuseWidth)) {
    if (Math.abs(vtY) > fuseWidth / 2 + VALUE_TOL) {
      vtMountedOffFuselage = true;
      feedback.push(STRINGS.attachment.vtWing);
    }
  } else {
    feedback.push("Unable to verify vertical tail lateral placement due to missing geometry data");
    controlFailures += 1;
  }

  const strakeArea = asNumber(getCell(main, "D18"));
  if (Number.isFinite(strakeArea) && strakeArea > 1) {
    const sweep = asNumber(getCell(geom, "K15"));
    const y = asNumber(getCell(geom, "M152"));
    const strake = asNumber(getCell(geom, "L155"));
    const apex = asNumber(getCell(geom, "L38"));
    if (
      Number.isFinite(sweep) &&
      Number.isFinite(y) &&
      Number.isFinite(strake) &&
      Number.isFinite(apex)
    ) {
      const wing = y / Math.tan((90 - sweep) * DEG_TO_RAD) + apex;
      if (!(wing < strake + 0.5)) {
        feedback.push(STRINGS.attachment.strake);
        controlFailures += 1;
      }
    } else {
      feedback.push("Unable to verify strake attachment due to missing geometry data");
      controlFailures += 1;
    }
  }

  if (Number.isFinite(fuselageLength)) {
    const componentPositions = [];
    for (let col = 2; col <= 8; col += 1) {
      const val = asNumber(getCellByIndex(main, 23, col));
      if (Number.isFinite(val)) {
        componentPositions.push(val);
      }
    }
    if (componentPositions.some((value) => value >= fuselageLength)) {
      feedback.push(format(STRINGS.attachment.fuselage, fuselageLength));
      controlFailures += 1;
    }
  }

  if (vtMountedOffFuselage) {
    const vtApex = getPlanformPoint(geom, 163, { absoluteY: false });
    const vtRootTE = getPlanformPoint(geom, 166, { absoluteY: false });
    const wingTE = getPlanformPoint(geom, 41, { absoluteY: false });
    if (
      !Number.isFinite(vtApex.x) ||
      !Number.isFinite(vtRootTE.x) ||
      !Number.isFinite(wingTE.x)
    ) {
      feedback.push("Unable to verify vertical tail overlap with wing due to missing geometry data");
      controlFailures += 1;
    } else {
      const chord = vtRootTE.x - vtApex.x;
      const overlap = Math.max(0, Math.min(wingTE.x, vtRootTE.x) - vtApex.x);
      if (!(chord > 0) || overlap + VALUE_TOL < VT_WING_FRACTION * chord) {
        feedback.push(STRINGS.attachment.vtOverlap);
        controlFailures += 1;
      }
    }
  }

  const wingAR = asNumber(getCell(main, "B19"));
  const pcsAR = asNumber(getCell(main, "C19"));
  const vtAR = asNumber(getCell(main, "H19"));
  if (Number.isFinite(wingAR) && Number.isFinite(pcsAR) && pcsAR > wingAR + AR_TOL) {
    feedback.push(format(STRINGS.attachment.aspectRatioPcs, pcsAR, wingAR));
    controlFailures += 1;
  }
  if (Number.isFinite(wingAR) && Number.isFinite(vtAR) && vtAR >= wingAR - AR_TOL) {
    feedback.push(format(STRINGS.attachment.aspectRatioVt, vtAR, wingAR));
    controlFailures += 1;
  }

  const engineDiameter = asNumber(getCell(main, "H29"));
  const inletX = asNumber(getCell(main, "F31"));
  const compressorX = asNumber(getCell(main, "F32"));
  const engineStartX =
    Number.isFinite(inletX) && Number.isFinite(compressorX) ? inletX + compressorX : Number.NaN;

  const widthValues = [];
  for (let row = 34; row <= 53; row += 1) {
    const stationX = asNumber(getCell(main, `B${row}`));
    const width = asNumber(getCell(main, `E${row}`));
    if (Number.isFinite(width) && Number.isFinite(stationX) && Number.isFinite(engineStartX) && stationX >= engineStartX) {
      widthValues.push(width);
    }
  }
  const minWidth = widthValues.length > 0 ? Math.min(...widthValues) : Number.NaN;
  const maxWidth = widthValues.length > 0 ? Math.max(...widthValues) : Number.NaN;

  if (Number.isFinite(engineDiameter) && Number.isFinite(minWidth)) {
    const requiredWidth = engineDiameter + 0.5;
    if (minWidth + VALUE_TOL <= requiredWidth) {
      feedback.push(format(STRINGS.attachment.fuselageWidth, minWidth, requiredWidth));
      controlFailures += 1;
    }
  } else {
    feedback.push("Unable to verify fuselage width clearance for engines");
    controlFailures += 1;
  }

  const allowedOverhang = Number.isFinite(maxWidth) ? 2.5 * maxWidth : Number.NaN;
  if (!Number.isFinite(allowedOverhang)) {
    feedback.push("Unable to compute fuselage width limit for aft overhang checks");
    controlFailures += 1;
  } else if (Number.isFinite(fuselageLength)) {
    const pcsTipX = getMaxX(geom, [117, 118]);
    const vtTipX = getMaxX(geom, [165, 166]);
    [
      { label: STRINGS.attachment.surfaceNames.pcs, tipX: pcsTipX },
      { label: STRINGS.attachment.surfaceNames.vt, tipX: vtTipX },
    ].forEach(({ label, tipX }) => {
      if (!Number.isFinite(tipX)) {
        return;
      }
      const overhang = tipX - fuselageLength;
      if (overhang > allowedOverhang + VALUE_TOL) {
        feedback.push(format(STRINGS.attachment.tipOverhang, label, overhang, allowedOverhang));
        controlFailures += 1;
      }
    });
  }

  if (Number.isFinite(engineDiameter)) {
    const engineLength = asNumber(getCell(main, "I29"));
    if (
      Number.isFinite(fuselageLength) &&
      Number.isFinite(inletX) &&
      Number.isFinite(compressorX) &&
      Number.isFinite(engineLength)
    ) {
      const protrusion = inletX + compressorX + engineLength - fuselageLength;
      if (protrusion > engineDiameter + VALUE_TOL) {
        feedback.push(format(STRINGS.attachment.engineProtrusion, protrusion, engineDiameter));
        controlFailures += 1;
      }
    } else {
      feedback.push("Unable to verify engine protrusion due to missing geometry data");
      controlFailures += 1;
    }
  }

  if (controlFailures > 0) {
    const deduction = Math.min(2, controlFailures);
    feedback.push(format(STRINGS.attachment.deduction, deduction));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
