export function reproducibleBuildTimestamp(env = process.env, now = () => new Date()) {
  const sourceDateEpoch = env.SOURCE_DATE_EPOCH;
  if (sourceDateEpoch === undefined || sourceDateEpoch === '') {
    return now().toISOString();
  }

  if (!/^\d+$/.test(sourceDateEpoch)) {
    throw new Error('SOURCE_DATE_EPOCH must be a non-negative integer number of seconds');
  }
  const milliseconds = Number(sourceDateEpoch) * 1000;
  if (!Number.isSafeInteger(milliseconds)) {
    throw new Error('SOURCE_DATE_EPOCH is outside the supported JavaScript date range');
  }
  const timestamp = new Date(milliseconds);
  if (Number.isNaN(timestamp.valueOf())) {
    throw new Error('SOURCE_DATE_EPOCH is outside the supported JavaScript date range');
  }
  return timestamp.toISOString();
}
