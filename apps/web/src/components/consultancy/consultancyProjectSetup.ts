const INVALID_FOLDER_CHARACTERS = /[<>:"/\\|?*]/g;

export function projectFolderName(name: string, date: Date): string {
  const safeName = name
    .trim()
    .replace(INVALID_FOLDER_CHARACTERS, "-")
    .replace(/\p{Cc}/gu, "-")
    .replace(/\s+/g, " ")
    .slice(0, 80)
    .replace(/[. ]+$/, "");
  if (!safeName || safeName === "." || safeName === "..") {
    throw new Error("Give the project a name that can be used as a folder name.");
  }
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${safeName}_${year}-${month}-${day}`;
}

export function projectFolderPath(parent: string, name: string, date: Date): string {
  const selected = parent.trim();
  const root =
    selected === "/" || /^[A-Za-z]:[\\/]$/.test(selected)
      ? selected
      : selected.replace(/[\\/]+$/, "");
  if (!root) throw new Error("Choose a parent folder for the project.");
  const separator = root.includes("\\") ? "\\" : "/";
  return `${root}${/[\\/]$/.test(root) ? "" : separator}${projectFolderName(name, date)}`;
}

export const CONSULTANCY_FOLDERS = [
  "bronnen",
  "word",
  "excel",
  "powerpoints",
  "oplevering",
  ".werkbestanden",
] as const;
