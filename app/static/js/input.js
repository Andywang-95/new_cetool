import Alpine from "alpinejs";
import { logStore } from "./func/logStore.js";
import windowApi from "./func/windowApi.js";

window.Alpine = Alpine;

Alpine.store("logStore", logStore);

Alpine.data("initData", () => ({
  tab: "review",
  reviewMode: "BOM_TipTop_PTC",
  importMode: "BOM_TipTop_PTC",
  showModal: false,
  showSettingModal: false,
  reviewBomPath: "",
  importBomPath: "",
  custom: { col: "A", row: "2" },
  tempCustom: { col: "A", row: "2" },
  settings: {},
  tempSettings: {},
  async init() {
    const resp = await fetch("/api/settings");
    const data = await resp.json();
    this.settings = data;
    this.tempSettings = data;
    this.$watch("$store.logStore.reviewLogs", () => {
      this.$refs.reviewLog.scrollTop = this.$refs.reviewLog.scrollHeight;
    });
    this.$watch("$store.logStore.importLogs", () => {
      this.$refs.importLog.scrollTop = this.$refs.importLog.scrollHeight;
    });
    this.$watch("$store.logStore.updateLogs", () => {
      this.$refs.updateLog.scrollTop = this.$refs.updateLog.scrollHeight;
    });
  },
  ...windowApi(),
}));
Alpine.start();
