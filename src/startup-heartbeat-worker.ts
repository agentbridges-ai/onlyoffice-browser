import { OFFICE_HOST_PROTOCOL, type OfficeHostStartupPhase } from './lib/office-host-protocol';

type StartMessage = {
  type: 'START';
  sessionId: string;
  phase: OfficeHostStartupPhase;
  intervalMs: number;
  port: MessagePort;
};

type PhaseMessage = {
  type: 'PHASE';
  phase: OfficeHostStartupPhase;
};

type StopMessage = { type: 'STOP' };

const scope = globalThis as unknown as {
  onmessage: ((event: MessageEvent<StartMessage | PhaseMessage | StopMessage>) => void) | null;
  close(): void;
};

let heartbeatPort: MessagePort | null = null;
let heartbeatTimer: ReturnType<typeof setInterval> | null = null;
let sessionId = '';
let phase: OfficeHostStartupPhase = 'connected';

function stop(): void {
  if (heartbeatTimer !== null) clearInterval(heartbeatTimer);
  heartbeatTimer = null;
  heartbeatPort?.close();
  heartbeatPort = null;
}

function heartbeat(): void {
  heartbeatPort?.postMessage({
    protocol: OFFICE_HOST_PROTOCOL,
    type: 'STARTUP_HEARTBEAT',
    sessionId,
    phase,
  });
}

scope.onmessage = (event) => {
  const message = event.data;
  if (message.type === 'PHASE') {
    phase = message.phase;
    heartbeat();
    return;
  }
  if (message.type === 'STOP') {
    stop();
    scope.close();
    return;
  }

  stop();
  sessionId = message.sessionId;
  phase = message.phase;
  heartbeatPort = message.port;
  heartbeatPort.start();
  heartbeat();
  heartbeatTimer = setInterval(heartbeat, message.intervalMs);
};
