const DEFAULT_CONFIG = {
  SPREADSHEET_ID: "1bTRSe5l7RTMk1taHNtYaAUMcFBEIwGUf6Yz0icPtp2M",
  LINE_WORKS_WEBHOOK_URL: "https://webhook.worksmobile.com/message/02c456e3-6055-494b-8fe1-026271215198",
  GPT_API_KEY: "",
};

function sanitizeSpreadsheetId(value) {
  if (!value) return "";
  const trimmed = String(value).trim();
  const editIndex = trimmed.indexOf('/edit');
  let base = editIndex !== -1 ? trimmed.slice(0, editIndex) : trimmed;
  const queryIndex = base.indexOf('?');
  if (queryIndex !== -1) {
    base = base.slice(0, queryIndex);
  }
  return base.trim();
}

function getConfigValue(key) {
  if (!key) {
    return "";
  }

  const props = PropertiesService.getScriptProperties();
  let value = props.getProperty(key) || DEFAULT_CONFIG[key] || "";

  if (key === 'SPREADSHEET_ID' && value) {
    const sanitized = sanitizeSpreadsheetId(value);
    if (sanitized) {
      if (props.getProperty(key) !== sanitized) {
        props.setProperty(key, sanitized);
      }
      value = sanitized;
    }
  }

  return value || "";
}

function setConfigValue(key, value) {
  if (!key) throw new Error('key は必須です');
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}
