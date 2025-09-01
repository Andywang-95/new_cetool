export default function windowApi() {
  return {
    saveSettings(settings) {
      if (window.pywebview && window.pywebview.api) {
        window.pywebview.api.save_settings(settings);
      } else {
        console.log("Mock saveSettings", settings);
      }
    },
    selectBOM() {
      if (window.pywebview && window.pywebview.api) {
        return window.pywebview.api.select_bom_path();
      } else {
        console.log("Mock selectBOM");
        return Promise.resolve("/Users/mock/path/to/file.xlsx");
      }
    },
    // addLog(type, msg) {
    //   if (type === "review") {
    //     this.reviewLogs.push(msg);
    //   } else if (type === "import") {
    //     this.importLogs.push(msg);
    //   } else if (type === "update") {
    //     this.updateLogs.push(msg);
    //   }
    // },
    runReview() {
      if (window.pywebview && window.pywebview.api) {
        console.log("Running review...");
        return window.pywebview.api.run_review(
          this.selectedMode,
          this.reviewBomPath
        );
      } else {
        console.log("Mock runReview", this.selectedMode, this.reviewBomPath);
        return;
      }
    },
  };
}
