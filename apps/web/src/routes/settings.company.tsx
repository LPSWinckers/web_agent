import { createFileRoute } from "@tanstack/react-router";

import { CompanySettingsPanel } from "../components/settings/CompanySettings";

export const Route = createFileRoute("/settings/company")({ component: CompanySettingsPanel });
