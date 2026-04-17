import { getCell, asNumber } from "./parseUtils.js";

const X_COLUMN = "L";
const LEFT_Y_COLUMN = "M";
const RIGHT_Y_COLUMN = "N";

const DEG_PER_RAD = 180 / Math.PI;

function normalizeValue(value) {
  return Number.isFinite(value) ? value : Number.NaN;
}

function buildRef(column, row) {
  return `${column}${row}`;
}

export function getPlanformPoint(geom, row, { absoluteY = true } = {}) {
  const x = asNumber(getCell(geom, buildRef(X_COLUMN, row)));
  const leftY = asNumber(getCell(geom, buildRef(LEFT_Y_COLUMN, row)));
  const rightY = asNumber(getCell(geom, buildRef(RIGHT_Y_COLUMN, row)));

  let y = Number.isFinite(rightY) ? rightY : leftY;
  if (!Number.isFinite(y)) {
    y = 0;
  }

  return {
    x: normalizeValue(x),
    y: absoluteY ? Math.abs(y) : normalizeValue(y),
  };
}

export function computeEdgeAngle(geom, startRow, endRow) {
  const start = getPlanformPoint(geom, startRow);
  const end = getPlanformPoint(geom, endRow);
  if (!Number.isFinite(start.x) || !Number.isFinite(end.x)) {
    return Number.NaN;
  }
  const dx = Math.abs(end.x - start.x);
  const dy = Math.abs(end.y - start.y);
  if (dx === 0 && dy === 0) {
    return 0;
  }
  return Math.atan2(dy, dx) * DEG_PER_RAD;
}

export function computeWingTrailingPlanformAngle(geom) {
  return computeEdgeAngle(geom, 40, 41);
}

export function computePcsTrailingPlanformAngle(geom) {
  return computeEdgeAngle(geom, 117, 118);
}

export function getChordLength(geom, leadingRow, trailingRow) {
  const leading = getPlanformPoint(geom, leadingRow, { absoluteY: false });
  const trailing = getPlanformPoint(geom, trailingRow, { absoluteY: false });
  if (!Number.isFinite(leading.x) || !Number.isFinite(trailing.x)) {
    return Number.NaN;
  }
  return trailing.x - leading.x;
}

export function getMaxX(geom, rows) {
  let maxX = Number.NEGATIVE_INFINITY;
  rows.forEach((row) => {
    const point = getPlanformPoint(geom, row, { absoluteY: false });
    if (Number.isFinite(point.x)) {
      maxX = Math.max(maxX, point.x);
    }
  });
  return maxX;
}

export function getPointPair(geom, rowA, rowB) {
  return {
    a: getPlanformPoint(geom, rowA),
    b: getPlanformPoint(geom, rowB),
  };
}
