#!/usr/bin/env node

export const RELEASE_STATES = Object.freeze(['candidate', 'staged', 'verified', 'canary', 'promoted', 'post-verified']);

const TRANSITIONS = new Map([
  ['candidate', new Set(['staged'])],
  ['staged', new Set(['verified'])],
  ['verified', new Set(['canary'])],
  ['canary', new Set(['promoted'])],
  ['promoted', new Set(['post-verified'])],
]);

export function assertReleaseState(state) {
  if (!RELEASE_STATES.includes(state)) throw new Error(`unknown release state: ${state}`);
  return state;
}

export function assertReleaseTransition(from, to) {
  assertReleaseState(from);
  assertReleaseState(to);
  if (!TRANSITIONS.get(from)?.has(to)) throw new Error(`invalid release transition: ${from} -> ${to}`);
  return to;
}

export function assertReleaseStateHistory(history) {
  if (!Array.isArray(history) || history.length === 0) throw new TypeError('release state history must be non-empty');
  history.forEach(assertReleaseState);
  for (let index = 1; index < history.length; index += 1) {
    assertReleaseTransition(history[index - 1], history[index]);
  }
  return history;
}

export function releaseStateEvidence(history, evidence = {}) {
  assertReleaseStateHistory(history);
  return Object.freeze({ version: 1, history: [...history], state: history.at(-1), ...evidence });
}
