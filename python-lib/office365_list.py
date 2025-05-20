from office365_commons import get_sharepoint_type_descriptor
from dss_constants import DSSConstants


class Office365List(object):
    def __init__(self, parent, list_id):
        self.session = parent.session
        self.list_id = list_id
        self.parent = parent

    def get_columns(self):
        url = self.get_column_url()
        return self.session.get_all_items(url=url)

    def get_column_url(self):
        url = "/".join(
            [
                self.parent.get_site_url(), "lists/{}/columns".format(
                    self.list_id
                )
            ]
        )
        return url

    def get_next_row(self, filter={}):
        params = {"expand": "field"}
        if filter:
            params.update({"$filter": filter})
        url = self.get_next_list_row_url()
        for row in self.session.get_next_item(
            url=url,
            params=params,
            force_no_batch=True
        ):
            yield row

    def get_next_list_row_url(self):
        url = "/".join(
            [
                self.parent.get_site_url(),
                "lists/{}/items".format(
                    self.list_id
                )
            ]
        )
        return url

    def get_next_list_url(self):
        url = "/".join(
            [
                self.parent.get_site_url(),
                "lists/{}".format(
                    self.list_id
                )
            ]
        )
        return url

    def get_lists_url(self):
        url = "/".join(
            [
                self.parent.get_site_url(),
                "lists"
            ]
        )
        return url

    def add_column(self, name, type, description=None):
        description = "" or description
        data = {
            "description": description,
            "enforceUniqueValues": False,
            "hidden": False,
            "indexed": False,
            "name": name,
        }
        data.update(
            get_sharepoint_type_descriptor(type)
        )
        url = self.get_column_url()
        self.session.request(
            method="POST",
            url=url,
            headers=DSSConstants.JSON_HEADERS,
            json=data,
            raise_on={403: "Check that your Azure app has Sites.Manage.All scope enabled"}
        )

    def write_row(self, row):
        headers = DSSConstants.JSON_HEADERS
        data = {
            "fields": row,
        }
        url = self.get_next_list_row_url()
        self.session.request(
            method="POST",
            url=url,
            headers=headers,
            json=data
        )

    def delete_row(self, row_id):
        self.session.request(
            method="DELETE",
            url=self.get_list_row_id_url(row_id)
        )

    def get_list_row_id_url(self, row_id):
        url = "/".join(
            [
                self.parent.get_site_url(),
                "lists/{}/items/{}".format(
                    self.list_id,
                    row_id
                )
            ]
        )
        return url

    def delete_all_rows(self):
        self.session.start_batch_mode()
        for row in self.get_next_row():
            row_id = row.get("id")
            self.delete_row(row_id)
        self.session.close()

    def get_record_count(self):
        url = "/".join(
            [
                self.get_next_list_row_url(),
                "?expand=fields(select=ID)"
            ]
        )
        response = self.session.request(
            method="GET",
            url=url
        )
        json_response = None
        try:
            json_response = response.json()
        except Exception as er:
            print("ALX:error")
            print("ALX:error {} ".format(er))
        
        print("ALX:get_record_count:url={}, response={}".format(url, json_response))
