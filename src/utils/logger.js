/* global console */

const PREFIX = "[MailLighter]";

export const logger = {
  info(...args) {
    if (typeof console !== "undefined" && console.info) {
      console.info(PREFIX, ...args);
    }
  },
  warn(...args) {
    if (typeof console !== "undefined" && console.warn) {
      console.warn(PREFIX, ...args);
    }
  },
  error(...args) {
    if (typeof console !== "undefined" && console.error) {
      console.error(PREFIX, ...args);
    }
  },
};
