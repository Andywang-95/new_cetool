export const logStore = {
  reviewLogs: [],
  importLogs: [],
  updateLogs: [],
  addLog(type, msg) {
    this[`${type}Logs`].push(msg);
  },
};
