import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";

const LEG_COLUMNS = [11, 12, 13, 14, 16, 18, 19, 22, 23];
const ROWS = {
  altitude: 33,
  mach: 35,
  afterburner: 36,
  distance: 38,
  time: 39,
};
const ALT_TOL = 10;
const MACH_TOL = 0.05;
const TIME_TOL = 0.1;
const DIST_TOL = 0.5;

export function runMissionChecks(workbook) {
  const feedback = [];
  let errors = 0;

  const main = workbook.sheets.main;
  const constraintsMach = asNumber(getCell(main, "U4"));

  const readRowValues = (rowIndex) =>
    LEG_COLUMNS.map((col) => asNumber(getCellByIndex(main, rowIndex, col)));

  const altitude = readRowValues(ROWS.altitude);
  const mach = readRowValues(ROWS.mach);
  const afterburner = readRowValues(ROWS.afterburner);
  const distance = readRowValues(ROWS.distance);
  const time = readRowValues(ROWS.time);

  const pushIf = (condition, message) => {
    if (condition) {
      feedback.push(message);
      errors += 1;
    }
  };

  pushIf(Math.abs(altitude[0] - 0) > ALT_TOL || Math.abs(afterburner[0] - 100) > ALT_TOL, STRINGS.missionLegs[0]);
  pushIf(!(altitude[1] >= altitude[0] - ALT_TOL && altitude[1] <= altitude[2] + ALT_TOL), STRINGS.missionLegs[1]);
  pushIf(!(mach[1] >= mach[0] - MACH_TOL && mach[1] <= mach[2] + MACH_TOL), STRINGS.missionLegs[2]);
  pushIf(Math.abs(afterburner[1]) > ALT_TOL, STRINGS.missionLegs[3]);

  pushIf(
    altitude[2] < 35000 - ALT_TOL ||
      Math.abs(mach[2] - 0.9) > MACH_TOL ||
      Math.abs(afterburner[2]) > ALT_TOL,
    STRINGS.missionLegs[4]
  );
  pushIf(
    altitude[3] < 35000 - ALT_TOL ||
      Math.abs(mach[3] - 0.9) > MACH_TOL ||
      Math.abs(afterburner[3]) > ALT_TOL,
    STRINGS.missionLegs[5]
  );

  pushIf(
    altitude[4] < 35000 - ALT_TOL ||
      constraintsMach == null ||
      Math.abs(mach[4] - constraintsMach) > MACH_TOL ||
      Math.abs(afterburner[4]) > ALT_TOL ||
      distance[4] < 150 - DIST_TOL,
    STRINGS.missionLegs[6]
  );

  pushIf(
    altitude[5] < 30000 - ALT_TOL ||
      mach[5] < 1.2 - MACH_TOL ||
      Math.abs(afterburner[5] - 100) > ALT_TOL ||
      time[5] < 2 - TIME_TOL,
    STRINGS.missionLegs[7]
  );

  pushIf(
    altitude[6] < 35000 - ALT_TOL ||
      constraintsMach == null ||
      Math.abs(mach[6] - constraintsMach) > MACH_TOL ||
      Math.abs(afterburner[6]) > ALT_TOL ||
      distance[6] < 150 - DIST_TOL,
    STRINGS.missionLegs[8]
  );

  pushIf(
    altitude[7] < 35000 - ALT_TOL ||
      Math.abs(mach[7] - 0.9) > MACH_TOL ||
      Math.abs(afterburner[7]) > ALT_TOL,
    STRINGS.missionLegs[9]
  );
  pushIf(
    Math.abs(altitude[8] - 10000) > ALT_TOL ||
      Math.abs(mach[8] - 0.4) > MACH_TOL ||
      Math.abs(afterburner[8]) > ALT_TOL ||
      Math.abs(time[8] - 20) > TIME_TOL,
    STRINGS.missionLegs[10]
  );

  if (errors > 0) {
    const deduction = Math.min(2, errors);
    feedback.push(STRINGS.missionSummary.replace(\"%d\", deduction));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
