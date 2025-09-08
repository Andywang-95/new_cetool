from functools import partial

import pandas as pd

import app.services.utils as utils


class ReviewService:
    def __init__(self, data, bom_path, api):
        self.bom_path = bom_path
        self.database_path = data["database_path"]
        self.api = api
        self.log = partial(self.api.logs, "review")
        self.parent_dir, self.filename, self.filename_stem = utils.path_detail(bom_path)

    def run(self, col, row, method):
        """不同模式的共用流程"""
        msg = utils.check_bom(self.bom_path)
        if utils.check_and_log(msg, self.log):
            return

        # 檢查資料庫檔案狀態
        msg = utils.check_database(self.database_path)
        if utils.check_and_log(msg, self.log):
            return

        self.log(f"Starting review for 【{self.filename}】...")

        try:
            col = utils.columns_from_string(col)
            self._process(col, row, method)
        except KeyError as e:
            utils.review_other_logs(self.log, e=str(e))
            return
        # 重新讀取比對完成的 excel
        utils.hightlight_comment(
            f"{self.database_path}/mapping.xlsx", self.new_path, col, row
        )
        utils.review_other_logs(self.log, new_filename=self.new_filename)

    def _process(self, col, row, method):
        """不同 review 實作差異邏輯"""
        if method == "main":
            self._main_review(col, row)
        elif method == "system":
            self._system_review(col, row)
        elif method == "custom":
            self._custom_review(col, row)
        elif method == "result":
            self._result_review(col, row)

    def _main_review(self, col, row):
        # 讀出 BOM 和 mapping 檔案
        bom_df, mapping_comment, _ = utils.read_files(self.bom_path, self.database_path)
        # 複製前 5 行原始資料
        header_rows = bom_df.iloc[: row - 1].copy()
        header_rows.loc[header_rows.index[-1], header_rows.shape[1]] = "CE Comment"
        # 設定第 6 行是header，後續是資料內容，並新增 comment 欄位，
        bom_data = bom_df.iloc[row - 1 :].copy()
        bom_data.columns = bom_df.iloc[row - 2]
        bom_data["group"] = (bom_data["Action"] == "Add").cumsum()
        # 篩選不在 mapping 且長度為16的料號
        utils.find_unmatched(bom_data, mapping_comment, col, self.log)
        # 分配對應料號的原始comment
        bom_data["raw_comment"] = bom_data["Number"].map(mapping_comment)
        # 取得群組主料的comment
        main_comment = bom_data[bom_data["Action"] == "Add"].set_index("group")[
            "raw_comment"
        ]
        # 分配主料comment到所有替料
        bom_data["main_comment"] = bom_data["group"].map(main_comment)
        # 根據條件填入 CE Comment
        bom_data["CE Comment"] = bom_data.apply(
            utils.correct_comment,
            axis=1,
            method="main",
        )
        bom_data.drop(columns=["group", "raw_comment", "main_comment"], inplace=True)
        final_df = pd.concat(
            [header_rows, pd.DataFrame(bom_data.values)], ignore_index=True
        )
        self.new_path, self.new_filename = utils.save_to_excel(
            final_df, self.parent_dir, self.filename_stem
        )

    def _system_review(self, col, row):
        # 讀出BOM檔案
        bom_data, mapping_comment, _ = utils.read_files(
            self.bom_path, self.database_path
        )
        bom_data.columns = bom_data.iloc[0]
        bom_data["group"] = bom_data["主件料號"].notna().cumsum()
        # 篩選不在 mapping 且長度為16的料號
        utils.find_unmatched(bom_data, mapping_comment, col, self.log)
        # 分配對應料號的原始comment
        bom_data["raw_comment"] = bom_data["元件/替代料號"].map(mapping_comment)
        # 取得群組主料的comment
        main_comment = bom_data[bom_data["主件料號"].notna()].set_index("group")[
            "raw_comment"
        ]
        # 分配主料comment到所有替料
        bom_data["main_comment"] = bom_data["group"].map(main_comment)
        # 根據條件填入 CE Comment
        bom_data["CE Comment"] = bom_data.apply(
            utils.correct_comment, axis=1, method="system"
        )
        bom_data.drop(columns=["group", "raw_comment", "main_comment"], inplace=True)
        bom_data.iloc[0, bom_data.shape[1] - 1] = "CE Comment"
        self.new_path, self.new_filename = utils.save_to_excel(
            bom_data, self.parent_dir, self.filename_stem
        )

    def _custom_review(self, col, row):
        # 讀出BOM檔案
        bom_data, mapping_comment, _ = utils.read_files(
            self.bom_path, self.database_path
        )
        # 新增說明欄位
        new_col_idx = bom_data.shape[1]
        bom_data[new_col_idx] = ""
        # 根據第二欄料號填入對應值
        bom_data.iloc[row:, new_col_idx] = (
            bom_data.iloc[row:, col].map(mapping_comment).fillna("")
        )
        # 篩選不在 mapping 且長度為16的料號
        utils.find_unmatched(bom_data, mapping_comment, col, self.log)

        self.new_path, self.new_filename = utils.save_to_excel(
            bom_data, self.parent_dir, self.filename_stem
        )

    def _result_review(self, col, row):
        # 讀出BOM檔案
        bom_data, mapping_comment, _ = utils.read_files(
            self.bom_path, self.database_path
        )
        if "Total count:" not in str(bom_data.iloc[0, 0]):
            raise KeyError
        # 新增說明欄位
        new_col_idx = bom_data.shape[1]
        # 根據第二欄料號填入對應值
        bom_data[new_col_idx] = bom_data.iloc[:, col].map(mapping_comment).fillna("")
        bom_data.iloc[2, new_col_idx] = "CE Comment"
        # 篩選不在 mapping 且長度為16的料號
        utils.find_unmatched(bom_data, mapping_comment, col, self.log)

        self.new_path, self.new_filename = utils.save_to_excel(
            bom_data, self.parent_dir, self.filename_stem
        )


# def main_review_service(data, bom_path, api, main_col, main_row):
#     msg = utils.check_bom(bom_path)
#     if utils.check_and_log(msg, api):
#         return

#     # 檢查資料庫檔案狀態
#     msg = utils.check_database(data["database_path"])
#     if utils.check_and_log(msg, api):
#         return

#     parent_dir, filename, filename_stem = utils.path_detail(bom_path)

#     api.logs("review", f"Starting review for 【{filename}】...")

#     try:
#         # 讀出 BOM 和 mapping 檔案
#         bom_df, mapping_comment = utils.read_files(bom_path, data)
#         # 複製前 5 行原始資料
#         header_rows = bom_df.iloc[: main_row - 1].copy()
#         header_rows.loc[header_rows.index[-1], header_rows.shape[1]] = "CE Comment"
#         # 設定第 6 行是header，後續是資料內容，並新增 comment 欄位，
#         bom_data = bom_df.iloc[main_row - 1 :].copy()
#         bom_data.columns = bom_df.iloc[main_row - 2]
#         bom_data["group"] = (bom_data["Action"] == "Add").cumsum()
#         col = utils.columns_from_string(main_col)
#         # 篩選不在 mapping 且長度為16的料號
#         utils.find_unmatched(bom_data, mapping_comment, col, api)
#         # 分配對應料號的原始comment
#         bom_data["raw_comment"] = bom_data["Number"].map(mapping_comment)
#         # 取得群組主料的comment
#         main_comment = bom_data[bom_data["Action"] == "Add"].set_index("group")[
#             "raw_comment"
#         ]
#         # 分配主料comment到所有替料
#         bom_data["main_comment"] = bom_data["group"].map(main_comment)
#         # 根據條件填入 CE Comment
#         bom_data["CE Comment"] = bom_data.apply(
#             utils.correct_comment,
#             axis=1,
#             method="main",
#         )
#         bom_data.drop(columns=["group", "raw_comment", "main_comment"], inplace=True)
#         final_df = pd.concat(
#             [header_rows, pd.DataFrame(bom_data.values)], ignore_index=True
#         )
#         new_path, new_filename = utils.save_to_excel(
#             final_df, parent_dir, filename_stem
#         )
#     except KeyError as e:
#         utils.other_logs(api, e=str(e))
#         return

#     # 重新讀取比對完成的 excel
#     utils.hightlight_comment(
#         f"{data['database_path']}/mapping.xlsx",
#         new_path,
#         col,
#         main_row,
#     )

#     utils.other_logs(api, new_filename=new_filename)


# def result_review_service(data, bom_path, api, result_col, result_row):
#     msg = utils.check_bom(bom_path)
#     if utils.check_and_log(msg, api):
#         return

#     # 檢查資料庫檔案狀態
#     msg = utils.check_database(data["database_path"])
#     if utils.check_and_log(msg, api):
#         return

#     parent_dir, filename, filename_stem = utils.path_detail(bom_path)

#     api.logs("review", f"Starting review for 【{filename}】...")
#     try:
#         # 讀出BOM檔案
#         bom_data, mapping_comment = utils.read_files(bom_path, data)
#         if "Total count:" not in str(bom_data.iloc[0, 0]):
#             raise KeyError
#         # 新增說明欄位
#         new_col_idx = bom_data.shape[1]
#         # 根據第二欄料號填入對應值
#         col = utils.columns_from_string(result_col)
#         bom_data[new_col_idx] = bom_data.iloc[:, col].map(mapping_comment).fillna("")
#         bom_data.iloc[2, new_col_idx] = "CE Comment"
#         # 篩選不在 mapping 且長度為16的料號
#         utils.find_unmatched(bom_data, mapping_comment, col, api)

#         new_path, new_filename = utils.save_to_excel(
#             bom_data, parent_dir, filename_stem
#         )
#     except KeyError as e:
#         utils.other_logs(api, e=str(e))
#         return
#     # 重新讀取比對完成的 excel
#     utils.hightlight_comment(
#         f"{data['database_path']}/mapping.xlsx", new_path, col, result_row
#     )
#     utils.other_logs(api, new_filename=new_filename)


# def system_bom_review_service(data, bom_path, api, system_col, system_row):
#     msg = utils.check_bom(bom_path)
#     if utils.check_and_log(msg, api):
#         return

#     # 檢查資料庫檔案狀態
#     msg = utils.check_database(data["database_path"])
#     if utils.check_and_log(msg, api):
#         return

#     parent_dir, filename, filename_stem = utils.path_detail(bom_path)

#     api.logs("review", f"Starting review for 【{filename}】...")
#     try:
#         # 讀出BOM檔案
#         bom_data, mapping_comment = utils.read_files(bom_path, data)
#         bom_data.columns = bom_data.iloc[0]
#         bom_data["group"] = bom_data["主件料號"].notna().cumsum()
#         col = utils.columns_from_string(system_col)
#         # 篩選不在 mapping 且長度為16的料號
#         utils.find_unmatched(bom_data, mapping_comment, col, api)
#         # 分配對應料號的原始comment
#         bom_data["raw_comment"] = bom_data["元件/替代料號"].map(mapping_comment)
#         # 取得群組主料的comment
#         main_comment = bom_data[bom_data["主件料號"].notna()].set_index("group")[
#             "raw_comment"
#         ]
#         # 分配主料comment到所有替料
#         bom_data["main_comment"] = bom_data["group"].map(main_comment)
#         # 根據條件填入 CE Comment
#         bom_data["CE Comment"] = bom_data.apply(
#             utils.correct_comment, axis=1, method="system"
#         )
#         bom_data.drop(columns=["group", "raw_comment", "main_comment"], inplace=True)
#         bom_data.iloc[0, bom_data.shape[1] - 1] = "CE Comment"
#         new_path, new_filename = utils.save_to_excel(
#             bom_data, parent_dir, filename_stem
#         )
#     except KeyError as e:
#         utils.other_logs(api, e=str(e))
#         return

#     # 重新讀取比對完成的 excel
#     utils.hightlight_comment(
#         f"{data['database_path']}/mapping.xlsx",
#         new_path,
#         col,
#         system_row,
#     )

#     utils.other_logs(api, new_filename=new_filename)


# def custom_review_service(data, bom_path, api, cus_col, cus_row):
#     msg = utils.check_bom(bom_path)
#     if utils.check_and_log(msg, api):
#         return

#     # 檢查資料庫檔案狀態
#     msg = utils.check_database(data["database_path"])
#     if utils.check_and_log(msg, api):
#         return

#     parent_dir, filename, filename_stem = utils.path_detail(bom_path)

#     api.logs("review", f"Starting review for 【{filename}】...")
#     # 讀出BOM檔案
#     bom_data, mapping_comment = utils.read_files(bom_path, data)
#     # 新增說明欄位
#     new_col_idx = bom_data.shape[1]
#     bom_data[new_col_idx] = ""
#     # 根據第二欄料號填入對應值
#     col = utils.columns_from_string(cus_col)
#     bom_data.iloc[cus_row:, new_col_idx] = (
#         bom_data.iloc[cus_row:, col].map(mapping_comment).fillna("")
#     )
#     # 篩選不在 mapping 且長度為16的料號
#     utils.find_unmatched(bom_data, mapping_comment, col, api)

#     new_path, new_filename = utils.save_to_excel(bom_data, parent_dir, filename_stem)

#     # 重新讀取比對完成的 excel
#     utils.hightlight_comment(
#         f"{data['database_path']}/mapping.xlsx", new_path, col, cus_row
#     )

#     utils.other_logs(api, new_filename=new_filename)
