export const logStore = {
  reviewLogs: [],
  importLogs: [],
  updateLogs: [],
  addLog(type, msg) {
    if (type === "review") {
      this.reviewLogs.push(msg);
    } else if (type === "import") {
      this.importLogs.push(msg);
    } else if (type === "update") {
      this.updateLogs.push(msg);
    }
  },
};
