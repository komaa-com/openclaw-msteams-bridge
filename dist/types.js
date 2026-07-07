// Core call types for the self-contained Teams voice plugin.
// (CallRecord mirrors the shape voice-call persisted; trimmed to what a Teams realtime plugin needs.)
export const TERMINAL_STATES = new Set(["completed", "failed"]);
