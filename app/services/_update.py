import traceback
from copy import copy
from functools import partial

from tabulate import tabulate

import app.services.utils as utils


class UpdateService:
    def __init__(self, data, log_func):
        self.database_path = data["database_path"]
        self.log = partial(log_func, "update")

    def run(self):
        # 檢查資料庫檔案狀態
        msg = utils.check_database(self.database_path)
        if utils.check_and_log(msg, self.log):
            return
        self.log("Starting update...")
        try:
            self._process()
        except Exception as e:
            self.log(
                "Error:\n"
                + str(traceback.format_exc())
                + "\n----------------------------------------\n"
            )

    def _process(self):
        # 讀取 mapping.xlsx 與 maintain.xlsx
        (wb_mapping, ws_mapping), wb_maintain, mapping_path, maintain_path = (
            utils.read_db_files(self.database_path)
        )

        # 頁籤列表
        sheet_names = ["機構料件", "電子料(1)", "電子料(2)", "電子料(R,C)", "Others"]

        for name in sheet_names:
            upd = []
            new = []
            color = []
            ws_maintain = wb_maintain[name]
            dict_maintain = utils.to_dict(ws_maintain)
            dict_mapping = utils.to_dict(ws_mapping)
            for pn in dict_maintain:
                if pn in dict_mapping:
                    # 檢查 CE Comment 是否不同
                    if dict_mapping[pn][2].value != dict_maintain[pn][2].value:
                        upd.append(
                            [pn, dict_mapping[pn][2].value, dict_maintain[pn][2].value]
                        )
                        dict_mapping[pn][2].value = dict_maintain[pn][2].value
                    # 檢查顏色或字體
                    if (
                        dict_mapping[pn][2].fill.start_color.rgb
                        != dict_maintain[pn][2].fill.start_color.rgb
                    ):
                        dict_mapping[pn][2].fill = copy(dict_maintain[pn][2].fill)
                        color.append([pn, dict_maintain[pn][2].value])
                else:
                    ws_mapping.append(
                        [
                            pn,
                            dict_maintain[pn][0].value,
                            dict_maintain[pn][1].value,
                            dict_maintain[pn][2].value,
                        ]
                    )
                    new.append([pn, dict_maintain[pn][2].value])
            # log 格式化
            self.log(f"【{name}】：")
            if new:
                self.log(f"新增 {len(new)} 筆資料：")
                self.log(tabulate(new, headers=["PartNum", "Comment"], tablefmt="html"))
            if upd:
                self.log(f"\n更新comment {len(upd)} 筆資料：")
                self.log(
                    tabulate(
                        upd,
                        headers=["PartNum", "Old_Comment", "New_Comment"],
                        tablefmt="html",
                    )
                )
            if color:
                self.log(f"\n更新High Light底色 {len(color)} 筆資料")
                self.log(
                    tabulate(color, headers=["PartNum", "Comment"], tablefmt="html")
                )

        ws_mapping.protection.enable()
        # wb_mapping.save(mapping_path)
        wb_mapping.save(f"{self.database_path}/mapping.xlsx")
        wb_maintain.save(f"{self.database_path}/maintain.xlsx")
        wb_mapping.save(f"{utils.datetime.date.today()}mapping.xlsx")
        wb_maintain.save(f"{utils.datetime.date.today()}maintain.xlsx")
        self.log("\n更新完成\n----------------------------------------\n")
