export interface CanonicalBrokerAcceptanceDefaults {
  requiredLifecycleIterations: number;
  requiredConcurrentOrigins: number;
  maxBrokerReservedBytes: number;
  maxRecoveryMs: number;
  maxFinalHeapDeltaMb: number;
  maxFinalRssDeltaMb: number;
  stableWindowCycles: number;
  maxStableHeapRangeMb: number;
  maxStableRssRangeMb: number;
  maxStableHeapSlopeMbPerCycle: number;
  maxStableRssSlopeMbPerCycle: number;
}

export interface AcceptanceCheck {
  id: string;
  actual: unknown;
  expected: unknown;
  pass: boolean;
  detail: unknown;
}

export interface AcceptanceSection {
  pass: boolean;
  thresholds?: CanonicalBrokerAcceptanceDefaults;
  checks: AcceptanceCheck[];
  [key: string]: unknown;
}

export const CANONICAL_BROKER_ACCEPTANCE_DEFAULTS: Readonly<CanonicalBrokerAcceptanceDefaults>;
export function percentile(values: unknown[], percentileValue: number): number | null;
export function linearSlope(values: unknown[]): number | null;
export function analyzeLifecycleReport(
  report: Record<string, any>,
  overrides?: Partial<CanonicalBrokerAcceptanceDefaults>,
): AcceptanceSection;
export function analyzeStartupPerformance(report: Record<string, any>, baselineP95Ms: number): AcceptanceSection;
export function analyzeBrokerMetrics(
  input: Record<string, any>,
  overrides?: Partial<CanonicalBrokerAcceptanceDefaults>,
): AcceptanceSection;
export function analyzeRecoveryResults(
  results: Array<Record<string, any>>,
  overrides?: Partial<CanonicalBrokerAcceptanceDefaults>,
): AcceptanceSection;
export function analyzeReleaseIntegrity(input: Record<string, any>): AcceptanceSection;
export function analyzeSecurityProbes(input: Record<string, any>): AcceptanceSection;
export function analyzeTrafficEvidence(input: Record<string, any>): AcceptanceSection;
export function combineAcceptanceSections(sections: Record<string, AcceptanceSection>): {
  pass: boolean;
  sections: Record<string, AcceptanceSection>;
};
