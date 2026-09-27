import { describe, expect, it } from "vite-plus/test";
import { isConsultancyInternalPath } from "./consultancyData";
import { projectFolderName, projectFolderPath } from "./consultancyProjectSetup";

const date = new Date(2026, 8, 27);

describe("consultancy project folders", () => {
  it("creates a dated child folder on Windows and Unix paths", () => {
    expect(projectFolderPath("G:\\Clients\\Acme", "Market study", date)).toBe(
      "G:\\Clients\\Acme\\Market study_2026-09-27",
    );
    expect(projectFolderPath("/clients/acme/", "Market study", date)).toBe(
      "/clients/acme/Market study_2026-09-27",
    );
    expect(projectFolderPath("/", "Market study", date)).toBe("/Market study_2026-09-27");
  });

  it("keeps folder names valid on Windows without losing the project title", () => {
    expect(projectFolderName("  Strategy: Europe / US.  ", date)).toBe(
      "Strategy- Europe - US_2026-09-27",
    );
    expect(() => projectFolderName("...", date)).toThrow();
  });

  it("hides setup and working files while leaving deliverables visible", () => {
    expect(isConsultancyInternalPath(".werkbestanden/chart-data.csv")).toBe(true);
    expect(isConsultancyInternalPath("powerpoints/Presentatie.t3deck.json")).toBe(true);
    expect(isConsultancyInternalPath("bronnen/.keep")).toBe(true);
    expect(isConsultancyInternalPath("word/Projectbrief.docx")).toBe(false);
    expect(isConsultancyInternalPath("powerpoints/Presentatie.pptx")).toBe(false);
  });
});
