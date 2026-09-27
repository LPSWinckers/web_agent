import { CONSULTANCY_FOLDERS } from "./consultancyProjectSetup";

export interface ProjectSetupFile {
  readonly path: string;
  readonly contents: string;
}

export function consultancyProjectFolders(): ProjectSetupFile[] {
  return CONSULTANCY_FOLDERS.map((folder) => ({ path: `${folder}/.keep`, contents: "" }));
}
