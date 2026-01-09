export function pchip(xList, yList, xQuery) {
  const points = xList
    .map((x, idx) => ({ x: Number(x), y: Number(yList[idx]) }))
    .filter((p) => Number.isFinite(p.x) && Number.isFinite(p.y))
    .sort((a, b) => a.x - b.x);

  const unique = [];
  for (const point of points) {
    const last = unique[unique.length - 1];
    if (last && Math.abs(point.x - last.x) <= Number.EPSILON) {
      last.y = point.y;
    } else {
      unique.push({ ...point });
    }
  }

  const n = unique.length;
  if (n < 2 || !Number.isFinite(xQuery)) {
    return null;
  }

  const x = unique.map((p) => p.x);
  const y = unique.map((p) => p.y);
  const h = [];
  const delta = [];
  for (let i = 0; i < n - 1; i += 1) {
    const dx = x[i + 1] - x[i];
    if (dx === 0) {
      return null;
    }
    h.push(dx);
    delta.push((y[i + 1] - y[i]) / dx);
  }

  const m = new Array(n).fill(0);

  if (n === 2) {
    m[0] = delta[0];
    m[1] = delta[0];
  } else {
    const sign = (v) => (v === 0 ? 0 : v > 0 ? 1 : -1);
    const clampEndpoint = (value, delta0, delta1) => {
      if (!Number.isFinite(value) || delta0 === 0) {
        return 0;
      }
      if (sign(value) !== sign(delta0)) {
        return 0;
      }
      if (sign(delta0) !== sign(delta1) && Math.abs(value) > Math.abs(3 * delta0)) {
        return 3 * delta0;
      }
      return value;
    };

    const m0 = ((2 * h[0] + h[1]) * delta[0] - h[0] * delta[1]) / (h[0] + h[1]);
    const mn = ((2 * h[n - 2] + h[n - 3]) * delta[n - 2] - h[n - 2] * delta[n - 3]) / (h[n - 2] + h[n - 3]);
    m[0] = clampEndpoint(m0, delta[0], delta[1]);
    m[n - 1] = clampEndpoint(mn, delta[n - 2], delta[n - 3]);

    for (let i = 1; i < n - 1; i += 1) {
      if (delta[i - 1] === 0 || delta[i] === 0 || sign(delta[i - 1]) !== sign(delta[i])) {
        m[i] = 0;
      } else {
        const w1 = 2 * h[i] + h[i - 1];
        const w2 = h[i] + 2 * h[i - 1];
        m[i] = (w1 + w2) / (w1 / delta[i - 1] + w2 / delta[i]);
      }
    }
  }

  if (xQuery <= x[0]) {
    return y[0] + m[0] * (xQuery - x[0]);
  }
  if (xQuery >= x[n - 1]) {
    return y[n - 1] + m[n - 1] * (xQuery - x[n - 1]);
  }

  let idx = 0;
  while (idx < n - 2 && xQuery > x[idx + 1]) {
    idx += 1;
  }

  const t = (xQuery - x[idx]) / h[idx];
  const t2 = t * t;
  const t3 = t2 * t;

  const h00 = 2 * t3 - 3 * t2 + 1;
  const h10 = t3 - 2 * t2 + t;
  const h01 = -2 * t3 + 3 * t2;
  const h11 = t3 - t2;

  return h00 * y[idx] + h10 * h[idx] * m[idx] + h01 * y[idx + 1] + h11 * h[idx] * m[idx + 1];
}
