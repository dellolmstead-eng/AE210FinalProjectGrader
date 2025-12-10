import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

export function runFuelVolumeChecks(workbook) {
  const feedback = [];
  let delta = 0;

  const main = workbook.sheets.main;
  const fuelAvailable = asNumber(getCell(main, "O18"));
  const fuelRequired = asNumber(getCell(main, "X40"));
  if (Number.isFinite(fuelAvailable) && Number.isFinite(fuelRequired) && fuelAvailable < fuelRequired) {
    feedback.push(format(STRINGS.fuel.shortage, fuelAvailable, fuelRequired));
    delta -= 2;
  }

  const volumeRemaining = asNumber(getCell(main, "Q23"));
  if (!Number.isFinite(volumeRemaining) || volumeRemaining <= 0) {
    feedback.push(format(STRINGS.fuel.volume, volumeRemaining ?? 0));
    delta -= 2;
  }

  return { delta, feedback };
}
