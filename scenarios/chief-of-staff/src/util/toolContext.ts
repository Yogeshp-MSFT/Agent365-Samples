// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// Per-run tool context for deterministic flows that drive the LLM (Capture).
// Tools report their outcomes here so the caller can tell real success from a
// model summary that glosses over failed tool calls, and so a retried run can
// recognise work it already did. Scoped with AsyncLocalStorage, so concurrent
// user turns never see another run's context.

import { AsyncLocalStorage } from 'async_hooks';

export interface ToolOutcome {
  tool: string;
  ok: boolean;
  error?: string;
}

interface ToolRunContext {
  /** Stable identity of the run's source (e.g. `capture:<eventId>`). */
  sourceRef?: string;
  outcomes: ToolOutcome[];
}

const storage = new AsyncLocalStorage<ToolRunContext>();

export async function runWithToolContext<T>(
  sourceRef: string | undefined,
  fn: () => Promise<T>
): Promise<{ result: T; outcomes: ToolOutcome[] }> {
  const ctx: ToolRunContext = { sourceRef, outcomes: [] };
  const result = await storage.run(ctx, fn);
  return { result, outcomes: ctx.outcomes };
}

/** Source identity of the current run, if a deterministic flow started it. */
export function currentSourceRef(): string | undefined {
  return storage.getStore()?.sourceRef;
}

/** Record a tool outcome. No-op outside runWithToolContext. */
export function recordToolOutcome(outcome: ToolOutcome): void {
  storage.getStore()?.outcomes.push(outcome);
}
