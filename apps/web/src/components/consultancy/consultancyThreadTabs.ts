import type { ScopedThreadRef } from "@t3tools/contracts";

export type OpenConsultancyThreadDetail = {
  projectId: string;
  threadRef: ScopedThreadRef;
  title: string;
};

export type ShowConsultancyOverviewDetail = {
  environmentId: string;
  projectId: string;
};

const pendingThreadOpens: OpenConsultancyThreadDetail[] = [];

export function requestConsultancyThreadTab(detail: OpenConsultancyThreadDetail): void {
  pendingThreadOpens.push(detail);
  window.dispatchEvent(new Event("consultancy-open-thread"));
}

export function takeConsultancyThreadTabRequests(): OpenConsultancyThreadDetail[] {
  return pendingThreadOpens.splice(0);
}

export function requestConsultancyOverview(detail: ShowConsultancyOverviewDetail): void {
  window.dispatchEvent(new CustomEvent("consultancy-show-overview", { detail }));
}
