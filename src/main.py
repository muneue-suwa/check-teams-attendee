import csv
import re
from pathlib import Path
from tkinter import filedialog
from datetime import datetime
import argparse

import win32com.client


class CheckTeamsAttendee:
    def __init__(self):
        """初期化する
        """
        self.PROJ_DIRNAME = Path(__file__).resolve().parents[1]

        # ファイル名のデフォルト
        self.EXCEL_FILENAME = self.PROJ_DIRNAME / "名簿.xlsx"
        self.PASSWD_FILENAME = self.PROJ_DIRNAME / "password.txt"
        self.RESULT_FILENAME = self.PROJ_DIRNAME / "result.txt"

    def main(self):
        """Main script

        Raises:
            FileNotFoundError: 名簿のエクセルファイルが見つからなかったとき
            FileNotFoundError: パスワードファイルが見つからなかったとき
            FileNotFoundError: ファイルの選択がキャンセルされたとき
        """
        # Debug modeを判定する
        parser = argparse.ArgumentParser(
            description=(
                "Check Microsoft Teams meetings attendees "
                "using name list file"
            )
        )
        parser.add_argument(
            "namelist_filename",
            help=f"The name list filename. Default: {self.EXCEL_FILENAME}.",
            default=self.EXCEL_FILENAME, type=Path, nargs="?",
        )
        parser.add_argument(
            "--no-password",
            help="The name list file is not locked",
            action="store_true",
        )
        parser.add_argument(
            "--debug", help="Run program with debug mode",
            action="store_true",
        )
        args = parser.parse_args()
        self.EXCEL_FILENAME = args.namelist_filename
        HAS_NOT_PASSWORD = args.no_password
        IS_DEBUG_MODE = args.debug

        # 名簿とパスワードファイルの存在を確認する
        if not self.EXCEL_FILENAME.exists():
            raise FileNotFoundError(
                f"Download and save as {self.EXCEL_FILENAME}"
            )
        if not self.PASSWD_FILENAME.exists():
            raise FileNotFoundError(f"Create {self.PASSWD_FILENAME}")

        # 出席者リストのファイルを選択する
        # Debug modeのときは，./meetingAttendanceList.csvを用いる
        meeting_attendance_list_csv =\
            self.PROJ_DIRNAME / "meetingAttendanceList.csv"
        if IS_DEBUG_MODE is False:
            idir = "~/Downloads"
            filetype = [("出席者リスト", "*.csv")]
            temp_meeting_attendance_list_csv = filedialog.askopenfilename(
                filetypes=filetype, initialdir=idir
            )
            if temp_meeting_attendance_list_csv == "":
                raise FileNotFoundError("Canceled")
            meeting_attendance_list_csv = Path(
                temp_meeting_attendance_list_csv
            )

        # 出席者リストを取得する
        attendees_list = self.get_attendees_list_from_csv(
            meeting_attendance_list_csv
        )
        self.read_excel(attendees_list, HAS_NOT_PASSWORD)

    def get_attendees_list_from_csv(self, meeting_attendance_list_csv: Path):
        """出席者リストをCSVファイルから取得する

        Args:
            meeting_attendance_list_csv (Path): meetingAttendanceList.csvのパス

        Returns:
            list[str]: 出席者リスト
        """
        # 出席者リストのファイルからファイルを作成する
        attendees_list = []
        with meeting_attendance_list_csv.open(
            encoding="utf-16"
        ) as meeting_attendance_list_f:
            reader = csv.reader(meeting_attendance_list_f, delimiter="\t")
            for i, row in enumerate(reader):
                if i == 0:  # headerをパスする
                    pass
                temp_attendee_name = self.format_name(row[0])
                attendees_list.append(temp_attendee_name)
        return attendees_list

    def read_excel(self, attendees_list: list[str], has_not_password: bool):
        """名簿のエクセルファイルを読む

        Args:
            attendees_list (list[str]): 出席者リスト
            has_not_password (bool): パスワードを持っていない名簿
        """
        # Get password from password file
        if has_not_password is True:
            passwd = None
        else:
            with self.PASSWD_FILENAME.open(encoding="utf-8") as passwd_f:
                passwd = passwd_f.read().strip()

        # Excelファイルと出席者リストを比較し，未確認者の氏名とメールアドレスを取得する
        absentees_list = []
        try:
            excel = win32com.client.Dispatch('Excel.Application')
            workbook = excel.Workbooks.Open(
                self.EXCEL_FILENAME, False, False, None, passwd
            )
            worksheet = workbook.Worksheets[0]
            mail_list_str = ""
            for i in range(60):
                temp_name = worksheet.Cells.Item(i + 1, 2).Value
                if temp_name is None:
                    break
                temp_name = self.format_name(temp_name)
                do_attend = temp_name in attendees_list
                if do_attend is False:
                    # print(temp_name, do_attend)
                    temp_mail = worksheet.Cells.Item(i + 1, 4).Value
                    mail_list_str += f"{temp_mail},"
                    absentees_list.append(temp_name)
            if len(absentees_list) == 0:
                print("学生全員の出席が確認できました")
            else:
                self.export_result(absentees_list, mail_list_str)
        finally:
            excel.Quit()

    def export_result(
        self, absentees_list: list[str], mail_list_str: list[str]
    ):
        """欠席者リストを出力する

        Args:
            absentees_list (list[str]): 欠席者リスト
            mail_list_str (list[str]): 欠席者のメールアドレス
        """
        absentees_list_str = ""
        for absentee in absentees_list:
            absentees_list_str += f"{absentee.split('+',1)[0]}さん，"
        absentees_list_str = absentees_list_str[0:-1]
        teams_msg = f"現在，{absentees_list_str}の出席が確認できておりません"
        print(f"氏名|\n{absentees_list_str}")
        print(f"メアド|\n{mail_list_str}")
        print(f"Teams message|\n{teams_msg}")
        datetime_str = f"{datetime.now():%Y/%m/%d %H:%M:%S}"
        msg = f"{datetime_str}\n{absentees_list_str}\n"
        msg += f"{mail_list_str}\n{teams_msg}\n"
        with self.RESULT_FILENAME.open("w", encoding="utf-8") as result_f:
            result_f.write(msg)

    def format_name(self, raw_name: str):
        """氏名の空白文字の削除とアルファベットの大小文字の統一を行う．
        これは，Teams上で英字で氏名が登録されている人に対応するためのものである．

        Args:
            raw_name (str): もとの名前の文字列

        Returns:
            str: 整形された名前の文字列
        """
        # 先頭の1文字目を大文字，他を小文字に変換する
        formatted_name = raw_name.capitalize()
        # 空白文字を+に変換する
        formatted_name = re.sub(" |\u3000", "+", formatted_name)

        return formatted_name


if __name__ == "__main__":
    check_teams_attendee = CheckTeamsAttendee()
    check_teams_attendee.main()
