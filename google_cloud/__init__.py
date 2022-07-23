from __future__ import print_function
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from .actions import ActionChain
from .types import PasteType
from pandas import DataFrame
from typing import Union
import os



class GoogleSheet(object):
    ALL = "_ALL"

    def __init__(
        self,
        key_file: str,
        spreadsheet_id: str,
        api_name: str = "sheets",
        api_version: str = "v4",
        scope: list = ["https://www.googleapis.com/auth/spreadsheets"]
    ) -> None:
        assert os.path.isfile(key_file), '"%s" Key file does not exist' % key_file
        self._id = spreadsheet_id
        self._key = key_file
        self._credential = Credentials.from_service_account_file(self._key, scopes=scope)
        try:
            self._service = build(api_name, api_version, credentials=self._credential)
        except Exception:
            DISCOVERY_SERVICE_URL = "https://sheets.googleapis.com/$discovery/rest?version=v4"
            self._service = build(api_name, api_version, credentials=self._credential,
                                  discoveryServiceUrl=DISCOVERY_SERVICE_URL)
        self._spreadsheet = self._service.spreadsheets()
        info = self._spreadsheet.get(spreadsheetId=self._id).execute()
        self._sheets = {}
        for sheet in info["sheets"]:
            self._sheets[sheet["properties"]["title"]] = {
                "sheetId": sheet["properties"]["sheetId"],
                "index": sheet["properties"]["index"],
                "sheetType": sheet["properties"]["sheetType"],
                "rowCount": sheet["properties"]["gridProperties"]["rowCount"],
                "columnCount": sheet["properties"]["gridProperties"]["columnCount"]
            }
        self._url = info["spreadsheetUrl"]
        properties = info.get("properties", {})
        self._title = properties["title"]
        self._locale = properties["locale"]
        self._time_zone = properties["timeZone"]
        self._format = properties["defaultFormat"]
        self._setting = properties["spreadsheetTheme"]
        self._actions = ActionChain(self)

    @property
    def ID(self) -> int:
        return self._id

    @property
    def sheets(self) -> dict:
        return self._sheets

    @property
    def title(self) -> str:
        return self._title

    def get(self, sheet: str, name_box: str = ALL) -> DataFrame:
        if name_box == self.ALL:
            name_box = "A:" + ActionChain.count2str(self._sheets[sheet]["columnCount"])
        result = self._spreadsheet.values().get(
            spreadsheetId=self._id,
            range="%s!%s" % (sheet, name_box)
        ).execute()
        data = result.get("values", [])
        return DataFrame(data)

    def insert(self, sheet: str, name_box: str, values: Union[list[list], tuple[tuple], float, str, int]) -> dict:
        if isinstance(values, (float, int, str)):
            values = [[values]]
        return self._spreadsheet.values().update(
            spreadsheetId=self._id,
            range="%s!%s" % (sheet, name_box),
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()

    def insert_row(self, sheet: str, name_box: str, values: Union[list, tuple]) -> dict:
        return self.insert(sheet, name_box, [values])

    def insert_column(self, sheet: str, name_box: str, values: Union[list, tuple]) -> dict:
        return self.insert(sheet, name_box, [[v] for v in values])

    def clear(self, sheet: str) -> dict:
        return self._spreadsheet.values().clear(spreadsheetId=self._id, range=sheet).execute()

    def copy_paste(self, sheet: str, name_box_copy: str, name_box_paste: str, paste_type: str = PasteType.FORMULA) -> dict:
        self._actions.copy_paste(sheet, name_box_copy, name_box_paste, paste_type)
        return self._actions.perform()

    def cut_paste(self, sheet: str, name_box_cut: str, name_box_paste: str, paste_type: str = PasteType.FORMULA) -> dict:
        self._actions.cut_paste(sheet, name_box_cut, name_box_paste, paste_type)
        return self._actions.perform()
