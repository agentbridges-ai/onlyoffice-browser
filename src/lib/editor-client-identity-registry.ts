import {
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  type ResourceBrokerIdentity,
} from './resource-broker-protocol';

export type EditorClientIdentity = ResourceBrokerIdentity;

export type LiveEditorHostClient = {
  clientId: string;
  identity: EditorClientIdentity;
};

type ResolveEditorClientIdentityOptions = {
  clientId: string;
  resultingClientId?: string;
  releaseId: string | null;
  exactSourceIdentity?: EditorClientIdentity | null;
  connectedIdentity?: EditorClientIdentity | null;
  sourceIsOfficeRuntime: boolean;
  recoveryIdentities?: EditorClientIdentity[];
};

function isIdentity(value: EditorClientIdentity | null | undefined): value is EditorClientIdentity {
  return Boolean(value && isResourceBrokerReleaseId(value.releaseId) && isResourceBrokerSessionId(value.sessionId));
}

function sameIdentity(left: EditorClientIdentity, right: EditorClientIdentity): boolean {
  return left.releaseId === right.releaseId && left.sessionId === right.sessionId;
}

function matchesRelease(identity: EditorClientIdentity, releaseId: string | null): boolean {
  return releaseId === null || identity.releaseId === releaseId;
}

function uniqueRecoveryIdentity(candidates: readonly EditorClientIdentity[]): EditorClientIdentity | null {
  const unique = new Map<string, EditorClientIdentity>();
  for (const candidate of candidates) {
    if (!isIdentity(candidate)) continue;
    unique.set(`${candidate.releaseId}\n${candidate.sessionId}`, candidate);
  }
  return unique.size === 1 ? { ...(unique.values().next().value as EditorClientIdentity) } : null;
}

/**
 * Tracks the WindowClient lineage rooted at the exact versioned Office host.
 *
 * ONLYOFFICE creates same-origin nested editor frames. Their FetchEvents use
 * the nested WindowClient id rather than the outer office-host id, so the
 * Service Worker must propagate the session-bound identity across navigation
 * resultingClientIds. A new host binding clears the old lineage before an
 * origin slot can be reused.
 */
export class EditorClientIdentityRegistry {
  readonly #clients = new Map<string, EditorClientIdentity>();
  #boundIdentity: EditorClientIdentity | null = null;

  /**
   * Authorizes an exact Office Host source against the current WindowClient
   * snapshot. A different release/session cannot take over this origin while
   * any Host from the previous identity is still alive.
   *
   * The caller must obtain `liveHosts` from `clients.matchAll()` immediately
   * before binding and serialize authorization with the connection swap.
   */
  canBindHost(clientId: string, identity: EditorClientIdentity, liveHosts: readonly LiveEditorHostClient[]): boolean {
    if (!clientId || !isIdentity(identity)) return false;
    const live = this.#validateLiveHosts(liveHosts);
    if (!live) return false;
    const sourceIdentity = live.get(clientId);
    if (!sourceIdentity || !sameIdentity(sourceIdentity, identity)) return false;
    return [...live.values()].every((candidate) => sameIdentity(candidate, identity));
  }

  /**
   * A shell-prime client is not an Office Host. It may establish a transient
   * probe connection only while the origin has no live Office Host at all.
   */
  canBindPrime(identity: EditorClientIdentity, liveHosts: readonly LiveEditorHostClient[]): boolean {
    if (!isIdentity(identity)) return false;
    const live = this.#validateLiveHosts(liveHosts);
    return Boolean(live && live.size === 0);
  }

  bindHost(clientId: string, identity: EditorClientIdentity): void {
    if (!clientId || !isIdentity(identity)) throw new TypeError('Invalid editor host identity');
    if (this.#boundIdentity && !sameIdentity(this.#boundIdentity, identity)) {
      this.#clients.clear();
    }
    this.#boundIdentity = { ...identity };
    this.#clients.set(clientId, { ...identity });
  }

  resolve(options: ResolveEditorClientIdentityOptions): EditorClientIdentity | null {
    if (!options.clientId) return null;
    let identity = options.exactSourceIdentity;
    if (!isIdentity(identity)) identity = this.#clients.get(options.clientId) ?? null;

    if (!identity && options.sourceIsOfficeRuntime) {
      const recovered = uniqueRecoveryIdentity(options.recoveryIdentities ?? []);
      if (recovered) {
        identity =
          isIdentity(options.connectedIdentity) && sameIdentity(options.connectedIdentity, recovered)
            ? options.connectedIdentity
            : recovered;
      }
    }

    if (!isIdentity(identity) || !matchesRelease(identity, options.releaseId)) return null;
    this.#clients.set(options.clientId, { ...identity });
    if (options.resultingClientId) {
      this.#clients.set(options.resultingClientId, { ...identity });
    }
    return { ...identity };
  }

  #validateLiveHosts(liveHosts: readonly LiveEditorHostClient[]): Map<string, EditorClientIdentity> | null {
    const live = new Map<string, EditorClientIdentity>();
    for (const host of liveHosts) {
      if (!host || !host.clientId || !isIdentity(host.identity) || live.has(host.clientId)) return null;
      live.set(host.clientId, { ...host.identity });
    }
    return live;
  }
}
