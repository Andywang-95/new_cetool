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
    runReview() {
      if (window.pywebview && window.pywebview.api) {
        console.log("Running review...");
        return window.pywebview.api.run_review(
          this.reviewMode,
          this.reviewBomPath,
          this.custom.col,
          this.custom.row
        );
      } else {
        console.log("Mock runReview", this.reviewMode, this.reviewBomPath);
        return;
      }
    },
    runImport() {
      if (window.pywebview && window.pywebview.api) {
        console.log("Running import...");
        return window.pywebview.api.run_import(
          this.importMode,
          this.importBomPath
        );
      } else {
        console.log("Mock runImport", this.importMode, this.reviewBomPath);
        return;
      }
    },
    runUpdate() {
      if (window.pywebview && window.pywebview.api) {
        console.log("Running update...");
        return window.pywebview.api.run_update();
      } else {
        console.log("Mock runUpdate");
        return;
      }
    },
  };
}
