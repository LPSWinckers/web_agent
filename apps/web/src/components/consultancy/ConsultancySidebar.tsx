import { scopeProjectRef, scopeThreadRef } from "@t3tools/client-runtime/environment";
import { Link, useNavigate } from "@tanstack/react-router";
import {
  BriefcaseBusinessIcon,
  ChevronDownIcon,
  FileSpreadsheetIcon,
  PlusIcon,
  UserRoundIcon,
  XIcon,
} from "lucide-react";
import { useEffect, useMemo, useState } from "react";

import { useProjects, useThreadShells } from "~/state/entities";
import { useNewThreadHandler } from "~/hooks/useHandleNewThread";
import { Button } from "~/components/ui/button";
import { SidebarTrigger } from "~/components/ui/sidebar";
import { readCustomers, saveCustomers, splitCustomerProject } from "./consultancyData";
import { requestConsultancyThreadTab } from "./consultancyThreadTabs";

const ACTIVE_PROJECT_KEY = "t3-consultancy-active-project";
const PENDING_CUSTOMER_KEY = "t3-consultancy-pending-customer";

export function readPendingConsultancyCustomer(): string {
  return localStorage.getItem(PENDING_CUSTOMER_KEY) ?? "";
}

export function takePendingConsultancyCustomer(): string {
  const customer = readPendingConsultancyCustomer();
  localStorage.removeItem(PENDING_CUSTOMER_KEY);
  return customer;
}

export function selectPendingConsultancyCustomer(customer: string): void {
  localStorage.setItem(PENDING_CUSTOMER_KEY, customer);
  window.dispatchEvent(new Event("consultancy-customer-change"));
}

export function readActiveConsultancyProject(): string | null {
  return localStorage.getItem(ACTIVE_PROJECT_KEY);
}

export function selectConsultancyProject(projectId: string): void {
  localStorage.setItem(ACTIVE_PROJECT_KEY, projectId);
  window.dispatchEvent(new Event("consultancy-project-change"));
}

export function clearConsultancyProjectSelection(): void {
  localStorage.removeItem(ACTIVE_PROJECT_KEY);
  window.dispatchEvent(new Event("consultancy-project-change"));
}

export function ConsultancySidebar() {
  const navigate = useNavigate();
  const projects = useProjects();
  const threads = useThreadShells();
  const newThread = useNewThreadHandler();
  const [customers, setCustomers] = useState(readCustomers);
  const [expanded, setExpanded] = useState<string | null>(null);
  const [activeId, setActiveId] = useState(readActiveConsultancyProject);
  const [threadLimit, setThreadLimit] = useState(30);
  const [settingsOpen, setSettingsOpen] = useState(false);
  useEffect(() => {
    const update = () => setActiveId(readActiveConsultancyProject());
    window.addEventListener("consultancy-project-change", update);
    return () => window.removeEventListener("consultancy-project-change", update);
  }, []);
  const groups = useMemo(() => {
    const names = new Set(customers);
    for (const project of projects) {
      const entry = splitCustomerProject(project);
      if (entry) names.add(entry.customer);
    }
    return [...names].sort((a, b) => a.localeCompare(b));
  }, [customers, projects]);

  const addCustomer = () => {
    const name = window.prompt("Customer name")?.trim();
    if (!name || groups.some((customer) => customer.toLowerCase() === name.toLowerCase())) return;
    const next = [...customers, name];
    saveCustomers(next);
    selectPendingConsultancyCustomer(name);
    setCustomers(next);
    setExpanded(name);
    void navigate({ to: "/workbench" });
  };

  return (
    <aside className="flex h-full min-h-0 flex-col bg-sidebar text-sidebar-foreground">
      <div className="flex items-center justify-between border-b border-sidebar-border px-3 pb-4 pt-6">
        <div>
          <div className="flex items-center gap-2 text-sm font-semibold">
            <BriefcaseBusinessIcon className="size-4 text-primary" /> Consultancy
          </div>
          <p className="mt-1 text-xs text-muted-foreground">Customers and projects</p>
        </div>
        <SidebarTrigger aria-label="Collapse sidebar" />
      </div>
      <div className="flex items-center justify-between px-3 pb-2 pt-4">
        <span className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
          Customers
        </span>
        <Button size="icon-xs" variant="ghost" onClick={addCustomer} aria-label="Add customer">
          <PlusIcon />
        </Button>
      </div>
      <nav className="min-h-0 flex-1 overflow-y-auto px-2">
        {groups.length === 0 ? (
          <p className="px-2 py-4 text-xs text-muted-foreground">Add a customer to start.</p>
        ) : null}
        {groups.map((customer) => {
          const members = projects.filter(
            (project) => splitCustomerProject(project)?.customer === customer,
          );
          const open =
            expanded === customer ||
            (expanded === null && members.some((project) => project.id === activeId));
          return (
            <div key={customer} className="mb-1">
              <button
                type="button"
                className="flex w-full items-center gap-2 rounded-md px-2 py-2 text-left text-sm hover:bg-accent"
                onClick={() => setExpanded(open ? "" : customer)}
              >
                <ChevronDownIcon
                  className={`size-3.5 transition-transform ${open ? "" : "-rotate-90"}`}
                />
                <span className="min-w-0 flex-1 truncate">{customer}</span>
                <span className="text-xs text-muted-foreground">{members.length}</span>
              </button>
              {open && members.length === 0 ? (
                <button
                  type="button"
                  className="ml-7 flex items-center gap-1 px-2 py-1 text-xs text-muted-foreground hover:text-destructive"
                  onClick={() => {
                    const next = customers.filter((name) => name !== customer);
                    saveCustomers(next);
                    setCustomers(next);
                    setExpanded("");
                  }}
                >
                  <XIcon className="size-3" /> Remove empty customer
                </button>
              ) : null}
              {open ? (
                <div className="ml-5 border-l border-sidebar-border pl-2">
                  {members.map((project) => {
                    const projectThreads =
                      activeId === project.id
                        ? threads
                            .filter(
                              (thread) =>
                                thread.environmentId === project.environmentId &&
                                thread.projectId === project.id,
                            )
                            .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
                        : [];
                    return (
                      <div key={project.id}>
                        <button
                          type="button"
                          className={`block w-full truncate rounded-md px-2 py-1.5 text-left text-sm hover:bg-accent ${activeId === project.id ? "bg-accent font-medium" : ""}`}
                          onClick={() => {
                            selectConsultancyProject(project.id);
                            setActiveId(project.id);
                            setThreadLimit(30);
                            void navigate({ to: "/workbench" });
                          }}
                        >
                          {splitCustomerProject(project)?.name}
                        </button>
                        {activeId === project.id ? (
                          <div className="ml-3 border-l border-sidebar-border pl-2">
                            {projectThreads.slice(0, threadLimit).map((thread) => (
                              <button
                                key={thread.id}
                                type="button"
                                onClick={() => {
                                  if (readActiveConsultancyProject() !== project.id) {
                                    selectConsultancyProject(project.id);
                                  }
                                  void navigate({ to: "/workbench" }).then(() => {
                                    requestConsultancyThreadTab({
                                      projectId: project.id,
                                      threadRef: scopeThreadRef(thread.environmentId, thread.id),
                                      title: thread.title,
                                    });
                                  });
                                }}
                                className="block w-full truncate rounded-md px-2 py-1.5 text-left text-xs text-muted-foreground hover:bg-accent hover:text-foreground"
                              >
                                {thread.title}
                              </button>
                            ))}
                            {projectThreads.length > threadLimit ? (
                              <button
                                type="button"
                                className="px-2 py-1 text-xs text-primary"
                                onClick={() => setThreadLimit((limit) => limit + 30)}
                              >
                                Show more chats
                              </button>
                            ) : null}
                          </div>
                        ) : null}
                      </div>
                    );
                  })}
                  <Link
                    to="/workbench"
                    onClick={() => selectPendingConsultancyCustomer(customer)}
                    className="flex items-center gap-1 px-2 py-2 text-xs text-primary"
                  >
                    <PlusIcon className="size-3" /> New project
                  </Link>
                </div>
              ) : null}
            </div>
          );
        })}
      </nav>
      <div className="border-t border-sidebar-border p-3">
        <Button variant="ghost" className="w-full justify-start" render={<Link to="/workbench" />}>
          <FileSpreadsheetIcon /> Workspace
        </Button>
        {activeId ? (
          <Button
            variant="ghost"
            className="w-full justify-start"
            onClick={() => {
              const project = projects.find((entry) => entry.id === activeId);
              if (project) void newThread(scopeProjectRef(project.environmentId, project.id));
            }}
          >
            <PlusIcon /> New chat
          </Button>
        ) : null}
        <Button
          variant="ghost"
          className="w-full justify-start"
          aria-expanded={settingsOpen}
          aria-controls="workbench-settings-links"
          onClick={() => setSettingsOpen((open) => !open)}
        >
          <UserRoundIcon /> User settings
        </Button>
        {settingsOpen ? (
          <nav id="workbench-settings-links" aria-label="User settings" className="ml-4">
            <Button
              variant="ghost"
              className="w-full justify-start"
              render={<Link to="/settings/company" />}
            >
              Company settings
            </Button>
            <Button
              variant="ghost"
              className="w-full justify-start"
              render={<Link to="/settings/usage" />}
            >
              Usage
            </Button>
            <Button
              variant="ghost"
              className="w-full justify-start"
              render={<Link to="/settings/appearance" />}
            >
              Theme
            </Button>
          </nav>
        ) : null}
      </div>
    </aside>
  );
}
