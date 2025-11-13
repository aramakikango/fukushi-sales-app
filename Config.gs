const DEFAULT_CONFIG = {
  SPREADSHEET_ID: "",
  LINE_WORKS_WEBHOOK_URL: "",
};

function getConfigValue(key) {
  if (!key) {
    return "";
  }

  const stored = PropertiesService.getScriptProperties().getProperty(key);
  if (stored) {
    return stored;
  }

  return DEFAULT_CONFIG[key] || "";
}

function setConfigValue(key, value) {
  if (!key) throw new Error('key は必須です');
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}
