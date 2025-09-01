import Alpine from "alpinejs";
import { logStore } from "./func/logStore.js";
import windowApi from "./func/windowApi.js";

window.Alpine = Alpine;

Alpine.store("logStore", logStore);

Alpine.data("initData", () => ({
  tab: "review",
  selectedMode: "BOM_TipTop_PTC",
  showModal: false,
  showSettingModal: false,
  reviewBomPath: "",
  importBomPath: "",
  // reviewLogs: ["test review log 1", "test review log 2"],
  // importLogs: ["test import log 1", "test import log 2"],
  // updateLogs: ["test update log 1", "test update log 2"],
  settings: {},
  tempSettings: {},
  async init() {
    const resp = await fetch("/api/settings");
    const data = await resp.json();
    this.settings = data;
    this.tempSettings = data;
  },
  ...windowApi(),
}));
Alpine.start();
