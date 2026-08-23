export const PLUGIN_SPEC: string
export const SETTINGS_FILE_NAME: string
export const DATA_DIR_NAME: string
export const TOKENS_FILE_NAME: string
export const LOG_FILE_NAME: string

export function resolveConfigDir(explicit?: string): string

export interface SettingsResult {
  path: string
  created: boolean
  reason?: string
}

export interface InstallResult {
  changed: boolean
  reason?: string
  path: string
  backup?: string
  cacheLink?: string
  settings: SettingsResult
}

export function install(options?: {
  configDir?: string
  packageDir?: string
  force?: boolean
}): Promise<InstallResult>

export interface UninstallResult {
  changed: boolean
  reason?: string
  path?: string
  backup?: string
}

export function uninstall(options?: { configDir?: string }): Promise<UninstallResult>

export interface DoctorCheck {
  name: string
  ok: boolean
  detail?: string
}

export interface DoctorReport {
  configDir: string
  configPath: string
  checks: DoctorCheck[]
}

export function doctor(options?: { configDir?: string }): Promise<DoctorReport>
